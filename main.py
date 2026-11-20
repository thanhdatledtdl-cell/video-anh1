import os
import re
import sys
import json
import math
import time
import random
import asyncio
import subprocess
import urllib.request
import urllib.error
from xml.sax.saxutils import escape as xml_escape
from itertools import zip_longest

import gspread
from google.oauth2.service_account import Credentials as SACredentials
from google.oauth2.credentials import Credentials as OAuthCredentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
import edge_tts

# ==== CẤU HÌNH CHUNG ====
# ID mặc định của Sheet "Làm video ảnh 1" (dùng khi bấm Run workflow trực tiếp trên GitHub)
DEFAULT_SHEET_ID = (
    os.environ.get("SHEET_ID", "").strip()
    or "1EA0gsJYFutOI-UXgEdIhvwx90qdB1qGIuq6ukUEXqXA"
)
DEFAULT_DRIVE_OUTPUT_FOLDER_ID = os.environ.get("DRIVE_OUTPUT_FOLDER_ID", "").strip()
TEN_THU_MUC_GOC = "Video đầu ra"

GIONG_HOAI_MY = "vi-VN-HoaiMyNeural"
GIONG_NAM_MINH = "vi-VN-NamMinhNeural"
SO_LAN_THU_TTS = 3

AZURE_KEY = os.environ.get("AZURE_SPEECH_KEY", "").strip()
AZURE_REGION = os.environ.get("AZURE_SPEECH_REGION", "").strip()
PIPER_MODEL = "piper_models/vi.onnx"

VIENEU_BAT = os.environ.get("VIENEU", "1") != "0"
VIENEU_PRECISION = os.environ.get("VIENEU_PRECISION", "fp32")
_vieneu = None
VIENEU_HONG = False
EDGE_HONG = False
GIAN_CACH_TTS_GIAY = 2

LOGO_PATH = "logo.png"
TEXT_LIEN_HE = "Thành Đạt Led - 0986474671 -  0867933396"
TEXT_SDT = "0986474671 -  0867933396"
MAU_CHU = "#FF8A00"
THOI_GIAN_HIEN_TOI_THIEU_GIAY = 1.8
FONT_BOLD = "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"

GIOI_HAN_PHUT = 300
MAX_VIDEO = int(os.environ.get("MAX_VIDEO", "0") or 0)
MAX_ANH_MAC_DINH = 5

DONE = ("done", "hoàn thành", "hoan thanh")
S_NGAN = 720 / 1080

# ==== HIỆU ỨNG CHUYỂN CẢNH / CHUYỂN ĐỘNG ====
FPS = 30
# Mỗi "cảnh" (shot) dài khoảng 6-8 giây, ảnh sẽ được lặp lại nhiều lần với chuyển động khác nhau
SHOT_GIAY_MIN, SHOT_GIAY_MAX = 6.0, 8.0
MAX_SHOT = 50
# Thời gian 1 lần chuyển cảnh (giây), mỗi video chọn ngẫu nhiên trong khoảng này
CHUYEN_CANH_MIN, CHUYEN_CANH_MAX = 0.6, 1.0
HE_SO_ZOOM = 0.12  # phóng to tối đa 12% khi chuyển động

DS_CHUYEN_CANH = [
    "fade", "fadeblack", "fadewhite", "dissolve",
    "wipeleft", "wiperight", "wipeup", "wipedown",
    "slideleft", "slideright", "slideup", "slidedown",
    "smoothleft", "smoothright", "smoothup", "smoothdown",
    "circleopen", "circleclose", "circlecrop", "rectcrop", "radial",
    "vertopen", "vertclose", "horzopen", "horzclose",
    "diagtl", "diagtr", "diagbl", "diagbr",
    "hlslice", "hrslice", "vuslice", "vdslice",
    "squeezeh", "squeezev", "zoomin", "pixelize", "hblur",
]
KIEU_CHUYEN_DONG = ["in", "out", "trai_phai", "phai_trai", "tren_duoi", "duoi_tren"]
_XFADE_HOP_LE = None
_TRANG_THAI_HIEU_UNG = {"chuyen_cuoi": None}

# ==== MỖI TAB = 1 KIỂU VIDEO & THƯ MỤC RIÊNG TRÊN DRIVE ====
PROFILES = {
    "lam_video_dai": dict(
        W=1280, H=720,
        thu_muc_goc="Video Dai",
        voices=[GIONG_HOAI_MY, GIONG_NAM_MINH],
        vieneu_voice="Trúc Ly",
        logo_giua=False,
        logo_w=200, logo_y=20,
        contact_size=32, contact_y=50,
        sub_size=34, sub_margin_v=55, sub_margin_lr=80,
        so_tu_cum=11,
    ),
    "lam_video_ngan": dict(
        W=720, H=1280,
        thu_muc_goc="Video Ngan",
        voices=[GIONG_NAM_MINH, GIONG_HOAI_MY],
        vieneu_voice="Hải Đăng",
        khung_mo=True,
        # Khung ảnh vuông 1:1 (hiển thị 720x720 giữa video dọc), ảnh được giữ nguyên không cắt
        nen_w=1080, nen_h=1080,
        ten_text="Thành Đạt LED", lien_he_text=TEXT_SDT,
        mau_ten="white", mau_sdt="#FFD500",
        # Phụ đề nằm dưới khung ảnh, không che sản phẩm
        sub_size=36, sub_margin_v=110, sub_margin_lr=40,
        so_tu_cum=11,
    ),
}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

gc = drive_upload = spreadsheet = None
SHEET_ID = DEFAULT_SHEET_ID
DRIVE_OUTPUT_FOLDER_ID = DEFAULT_DRIVE_OUTPUT_FOLDER_ID
MAX_ANH = MAX_ANH_MAC_DINH


def lay_client_payload():
    p_path = os.environ.get("GITHUB_EVENT_PATH", "").strip()
    if p_path and os.path.exists(p_path):
        try:
            with open(p_path, "r", encoding="utf-8") as f:
                ev = json.load(f)
            cp = ev.get("client_payload") or {}
            if isinstance(cp, dict):
                return cp
        except Exception:
            pass
    return {}


def doc_cau_hinh_sheet():
    global DRIVE_OUTPUT_FOLDER_ID, MAX_ANH
    try:
        ws = spreadsheet.worksheet("CONFIG")
        vals = ws.get_all_values()
        cfg = {}
        for r in vals:
            if len(r) >= 2 and r[0].strip():
                cfg[r[0].strip()] = r[1].strip()
        if cfg.get("DRIVE_OUTPUT_FOLDER_ID"):
            DRIVE_OUTPUT_FOLDER_ID = cfg["DRIVE_OUTPUT_FOLDER_ID"]
        if cfg.get("MAX_ANH"):
            try:
                MAX_ANH = max(1, min(10, int(float(cfg["MAX_ANH"]))))
            except ValueError:
                MAX_ANH = MAX_ANH_MAC_DINH
        print(f"[CONFIG] Đã đọc cấu hình | FOLDER_ID={DRIVE_OUTPUT_FOLDER_ID or '(tự tạo)'} | MAX_ANH={MAX_ANH}", flush=True)
    except gspread.WorksheetNotFound:
        print("[CONFIG] Không có tab CONFIG, dùng cấu hình mặc định.", flush=True)


def luu_id_thu_muc_vao_config(folder_id):
    """Ghi (hoặc cập nhật) DRIVE_OUTPUT_FOLDER_ID vào tab CONFIG."""
    try:
        ws_cfg = spreadsheet.worksheet("CONFIG")
        vals = ws_cfg.get_all_values()
        for idx, r in enumerate(vals, start=1):
            if r and r[0].strip() == "DRIVE_OUTPUT_FOLDER_ID":
                ws_cfg.update_cell(idx, 2, folder_id)
                return
        ws_cfg.append_row(["DRIVE_OUTPUT_FOLDER_ID", folder_id, "Tự tạo bởi GitHub"])
    except Exception as e:
        print(f"[CONFIG] Không ghi được ID thư mục vào CONFIG: {e}", flush=True)


def khoi_tao():
    global gc, drive_upload, spreadsheet, SHEET_ID, DRIVE_OUTPUT_FOLDER_ID
    cp = lay_client_payload()
    if cp.get("sheet_id"):
        SHEET_ID = str(cp["sheet_id"]).strip()
    if not SHEET_ID:
        raise RuntimeError("Thiếu SHEET_ID.")

    sa_creds = SACredentials.from_service_account_file("service_account.json", scopes=SCOPES)
    gc = gspread.authorize(sa_creds)

    oauth_creds = OAuthCredentials(
        token=None,
        refresh_token=os.environ["OAUTH_REFRESH_TOKEN"],
        client_id=os.environ["OAUTH_CLIENT_ID"],
        client_secret=os.environ["OAUTH_CLIENT_SECRET"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    drive_upload = build("drive", "v3", credentials=oauth_creds)

    # In ra tài khoản Drive đang dùng để dễ kiểm tra
    try:
        u = retry(lambda: drive_upload.about().get(fields="user(emailAddress,displayName)").execute())["user"]
        print(f"[OAUTH] Đang dùng tài khoản Drive: {u.get('displayName')} <{u.get('emailAddress')}>", flush=True)
    except Exception as e:
        print(f"[OAUTH] Không đọc được tài khoản: {type(e).__name__}: {str(e)[:150]}", flush=True)

    spreadsheet = gc.open_by_key(SHEET_ID)
    doc_cau_hinh_sheet()

    # Kiểm tra thư mục gốc có truy cập được không, nếu không thì tự tạo thư mục mới
    can_luu_id = False
    if DRIVE_OUTPUT_FOLDER_ID:
        try:
            info = drive_upload.files().get(
                fileId=DRIVE_OUTPUT_FOLDER_ID, fields="id,name,trashed"
            ).execute()
            if info.get("trashed"):
                raise RuntimeError("thư mục đang ở thùng rác")
            print(f"[DRIVE] Dùng thư mục gốc: {info.get('name')} ({DRIVE_OUTPUT_FOLDER_ID})", flush=True)
        except Exception as e:
            print(f"[DRIVE] Tài khoản OAuth KHÔNG truy cập được thư mục {DRIVE_OUTPUT_FOLDER_ID} "
                  f"({type(e).__name__}: {str(e)[:120]}) -> tự tạo thư mục mới.", flush=True)
            DRIVE_OUTPUT_FOLDER_ID = ""

    if not DRIVE_OUTPUT_FOLDER_ID:
        DRIVE_OUTPUT_FOLDER_ID = retry(tim_hoac_tao_thu_muc, TEN_THU_MUC_GOC, "root", drive_upload)
        can_luu_id = True
        print(f"[DRIVE] Thư mục gốc mới: {TEN_THU_MUC_GOC} ({DRIVE_OUTPUT_FOLDER_ID})", flush=True)

    if can_luu_id:
        luu_id_thu_muc_vao_config(DRIVE_OUTPUT_FOLDER_ID)


# ================== TIỆN ÍCH ==================

class TTSBlocked(Exception):
    """Mọi nguồn giọng đọc đều không tạo được âm thanh."""


def retry(fn, *args, tries=5, **kwargs):
    for lan in range(tries):
        try:
            return fn(*args, **kwargs)
        except Exception as e:
            if lan == tries - 1:
                raise
            cho = 15 * (lan + 1)
            print(f"[RETRY {lan + 1}/{tries - 1}] {type(e).__name__}: {str(e)[:100]} - đợi {cho}s", flush=True)
            time.sleep(cho)


def ghi_log(ws, i, msg):
    try:
        retry(ws.batch_update, [
            {"range": f"F{i}", "values": [["❌ Lỗi"]]},
            {"range": f"I{i}", "values": [[msg]]},
        ], value_input_option="USER_ENTERED")
    except Exception as e:
        print(f"Không ghi được log vào Sheet: {e}", flush=True)


def slugify_vi(text):
    import unicodedata
    text = text.lower().replace("đ", "d")
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9]+", "-", text).strip("-") or "video"


def get_duration(path):
    r = subprocess.run(
        ["ffprobe", "-v", "error", "-show_entries", "format=duration", "-of", "json", path],
        capture_output=True, text=True,
    )
    return float(json.loads(r.stdout)["format"]["duration"])


def chay_ffmpeg(cmd):
    """Chạy ffmpeg, nếu lỗi thì báo kèm đoạn cuối thông báo lỗi của ffmpeg."""
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"ffmpeg lỗi: {(r.stderr or r.stdout)[-300:].strip()}")


def tim_folder(ten, parent_id, service):
    q = (f"name = '{ten.replace(chr(39), chr(92) + chr(39))}' "
         f"and mimeType = 'application/vnd.google-apps.folder' "
         f"and '{parent_id}' in parents and trashed = false")
    files = service.files().list(q=q, fields="files(id)").execute().get("files", [])
    return files[0]["id"] if files else None


def tim_hoac_tao_thu_muc(ten, parent_id, service):
    fid = tim_folder(ten, parent_id, service)
    if fid:
        return fid
    meta = {"name": ten, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
    return service.files().create(body=meta, fields="id").execute()["id"]


def tach_link_anh(raw_links):
    if not raw_links:
        return []
    parts = re.split(r"[\s,;]+", str(raw_links).strip())
    out, seen = [], set()
    for p in parts:
        p = p.strip()
        if p.startswith(("http://", "https://")) and p not in seen:
            seen.add(p)
            out.append(p)
            if len(out) >= MAX_ANH:
                break
    return out


def tai_anh_ve(url, out_path):
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                          "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp, open(out_path, "wb") as f:
        f.write(resp.read())
    if not os.path.exists(out_path) or os.path.getsize(out_path) < 500:
        raise RuntimeError(f"Ảnh tải về rỗng hoặc lỗi: {url}")


# ================== GIỌNG ĐỌC + PHỤ ĐỀ ==================

def giay_sang_ass_time(sec):
    h, m, s = int(sec // 3600), int((sec % 3600) // 60), int(sec % 60)
    return f"{h}:{m:02d}:{s:02d}.{int((sec - int(sec)) * 100):02d}"


def uoc_luong_tu(text, audio_path):
    dur = get_duration(audio_path)
    tu = text.split()
    if not tu:
        return []
    trong = [len(t) + 1 for t in tu]
    tong, t0, words = sum(trong), 0.0, []
    for t, w in zip(tu, trong):
        d = dur * w / tong
        words.append((t0, t0 + d, t))
        t0 += d
    return words


async def _edge_tts(text, audio_path, voice):
    communicate = edge_tts.Communicate(text, voice, boundary="WordBoundary")
    words = []
    with open(audio_path, "wb") as f:
        async for chunk in communicate.stream():
            if chunk["type"] == "audio":
                f.write(chunk["data"])
            elif chunk["type"] == "WordBoundary":
                s = chunk["offset"] / 10_000_000
                words.append((s, s + chunk["duration"] / 10_000_000, chunk["text"]))
    return words


def tao_giong_azure(text, audio_path, voice):
    ssml = ("<speak version='1.0' xml:lang='vi-VN'><voice name='%s'>%s</voice></speak>"
            % (voice, xml_escape(text)))
    req = urllib.request.Request(
        f"https://{AZURE_REGION}.tts.speech.microsoft.com/cognitiveservices/v1",
        data=ssml.encode("utf-8"),
        headers={
            "Ocp-Apim-Subscription-Key": AZURE_KEY,
            "Content-Type": "application/ssml+xml",
            "X-Microsoft-OutputFormat": "audio-24khz-96kbitrate-mono-mp3",
            "User-Agent": "video-automation",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=120) as r, open(audio_path, "wb") as f:
            f.write(r.read())
    except urllib.error.HTTPError as e:
        raise RuntimeError(f"Azure HTTP {e.code}: {e.read()[:150]!r}")
    if os.path.getsize(audio_path) < 1000:
        raise RuntimeError("Azure trả audio rỗng")
    return uoc_luong_tu(text, audio_path)


def tao_giong_vieneu(text, audio_path, voice_name):
    global _vieneu
    if _vieneu is None:
        from vieneu import Vieneu
        t0 = time.time()
        _vieneu = Vieneu(precision=VIENEU_PRECISION)
        print(f"[VieNeu] đã nạp model sau {time.time() - t0:.0f}s (precision={VIENEU_PRECISION})", flush=True)
    base = audio_path.rsplit(".", 1)[0]
    wav_path = base + "_vieneu.wav"
    try:
        t0 = time.time()
        audio = _vieneu.infer(text, voice=voice_name)
        _vieneu.save(audio, wav_path)
        subprocess.run(
            ["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", wav_path,
             "-ar", "44100", "-ac", "1", audio_path],
            check=True,
        )
        if os.path.getsize(audio_path) < 1000:
            raise RuntimeError("VieNeu trả audio rỗng")
        dur = get_duration(audio_path)
        print(f"[VieNeu] tạo {dur:.0f}s âm thanh mất {time.time() - t0:.0f}s (giọng {voice_name})", flush=True)
        return uoc_luong_tu(text, audio_path)
    finally:
        if os.path.exists(wav_path):
            os.remove(wav_path)


def tao_giong_piper(text, audio_path):
    if not os.path.exists(PIPER_MODEL):
        raise RuntimeError("chưa có model Piper")
    base = audio_path.rsplit(".", 1)[0]
    txt_path, wav_path = base + "_piper.txt", base + "_piper.wav"
    try:
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text)
        r = subprocess.run(
            [sys.executable, "-m", "piper", "-m", PIPER_MODEL, "-c", PIPER_MODEL + ".json",
             "-i", txt_path, "-f", wav_path],
            capture_output=True, text=True, timeout=900,
        )
        if r.returncode != 0 or not os.path.exists(wav_path):
            raise RuntimeError(f"Piper lỗi: {(r.stderr or r.stdout)[-200:]}")
        subprocess.run(
            ["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", wav_path,
             "-ar", "44100", "-ac", "1", audio_path],
            check=True,
        )
        if os.path.getsize(audio_path) < 1000:
            raise RuntimeError("Piper trả audio rỗng")
        return uoc_luong_tu(text, audio_path)
    finally:
        for p in (txt_path, wav_path):
            if os.path.exists(p):
                os.remove(p)


def tao_giong_doc(text, audio_path, voices, vieneu_voice="Hải Đăng"):
    global EDGE_HONG, VIENEU_HONG
    co_azure = bool(AZURE_KEY and AZURE_REGION)
    co_piper = os.path.exists(PIPER_MODEL)
    co_vieneu = VIENEU_BAT and not VIENEU_HONG
    loi = None

    if not EDGE_HONG:
        so_lan = 2 if (co_azure or co_piper or co_vieneu) else SO_LAN_THU_TTS
        for lan in range(so_lan):
            voice = voices[lan % len(voices)]
            try:
                time.sleep(GIAN_CACH_TTS_GIAY)
                words = asyncio.run(_edge_tts(text, audio_path, voice))
                if not os.path.exists(audio_path) or os.path.getsize(audio_path) < 1000:
                    raise RuntimeError("audio rỗng")
                if not words:
                    words = uoc_luong_tu(text, audio_path)
                return words
            except Exception as e:
                loi = e
                print(f"[TTS {lan + 1}/{so_lan}] edge {voice}: {type(e).__name__}: {str(e)[:120]}", flush=True)
                if isinstance(e, NameError):
                    break
                if lan < so_lan - 1:
                    time.sleep(10 * (lan + 1))
        EDGE_HONG = True
        print("[TTS] edge-tts đang hỏng, các video sau bỏ qua edge-tts.", flush=True)

    if co_azure:
        for voice in voices:
            try:
                words = tao_giong_azure(text, audio_path, voice)
                print(f"[TTS] dùng Azure Speech, giọng {voice}", flush=True)
                return words
            except Exception as e:
                loi = e
                print(f"[TTS Azure] {voice}: {str(e)[:150]}", flush=True)

    if co_vieneu:
        try:
            return tao_giong_vieneu(text, audio_path, vieneu_voice)
        except Exception as e:
            loi = e
            VIENEU_HONG = True
            print(f"[TTS VieNeu] {type(e).__name__}: {str(e)[:200]} -> bỏ qua VieNeu", flush=True)

    if co_piper:
        try:
            words = tao_giong_piper(text, audio_path)
            print("[TTS] dùng Piper (giọng offline miễn phí)", flush=True)
            return words
        except Exception as e:
            loi = e
            print(f"[TTS Piper] {type(e).__name__}: {str(e)[:200]}", flush=True)

    try:
        from gtts import gTTS
        gTTS(text=text, lang="vi").save(audio_path)
        if os.path.exists(audio_path) and os.path.getsize(audio_path) >= 1000:
            print("[TTS] dùng gTTS tiếng Việt", flush=True)
            return uoc_luong_tu(text, audio_path)
    except Exception as e:
        loi = e
        print(f"[TTS gTTS] {type(e).__name__}: {str(e)[:150]}", flush=True)

    raise TTSBlocked(f"mọi nguồn giọng đọc đều lỗi: {str(loi)[:120]}")


def viet_phu_de(words, ass_path, P):
    n = P["so_tu_cum"]
    cums = []
    for idx in range(0, len(words), n):
        g = words[idx: idx + n]
        cums.append((g[0][0], g[-1][1], " ".join(w[2] for w in g)))

    for idx, (start, end, nd) in enumerate(cums):
        gioi_han = cums[idx + 1][0] if idx + 1 < len(cums) else end + THOI_GIAN_HIEN_TOI_THIEU_GIAY
        cums[idx] = (start, max(min(start + THOI_GIAN_HIEN_TOI_THIEU_GIAY, gioi_han), end), nd)

    header = f"""[Script Info]
ScriptType: v4.00+
PlayResX: {P['W']}
PlayResY: {P['H']}
WrapStyle: 0

[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,DejaVu Sans,{P['sub_size']},&H00FFFFFF,&H00FFFFFF,&H00000000,&H00000000,-1,0,0,0,100,100,0,0,1,3,1,2,{P['sub_margin_lr']},{P['sub_margin_lr']},{P['sub_margin_v']},1

[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
"""
    lines = []
    for s, e, nd in cums:
        nd = nd.replace("\n", "\\N")
        lines.append(f"Dialogue: 0,{giay_sang_ass_time(s)},{giay_sang_ass_time(e)},Default,,0,0,0,,{nd}")
    with open(ass_path, "w", encoding="utf-8") as f:
        f.write(header + "\n".join(lines))
    print(f"[DEBUG] {len(words)} từ, {len(cums)} cụm phụ đề", flush=True)


# ================== DỰNG VIDEO NỀN: ẢNH VỪA KHUNG + CHUYỂN CẢNH ĐA DẠNG ==================

def ds_chuyen_canh_hop_le():
    """Chỉ giữ các kiểu chuyển cảnh mà bản ffmpeg đang cài hỗ trợ."""
    global _XFADE_HOP_LE
    if _XFADE_HOP_LE is None:
        ok = []
        try:
            r = subprocess.run(["ffmpeg", "-hide_banner", "-h", "filter=xfade"],
                               capture_output=True, text=True, timeout=20)
            txt = (r.stdout or "") + (r.stderr or "")
            ok = [t for t in DS_CHUYEN_CANH if re.search(r"\b" + t + r"\b", txt)]
        except Exception:
            ok = []
        if len(ok) < 4:
            ok = ["fade", "wipeleft", "wiperight", "slideleft", "slideright", "dissolve", "circleopen"]
        _XFADE_HOP_LE = ok
        print(f"[HIEU UNG] {len(ok)} kiểu chuyển cảnh khả dụng.", flush=True)
    return _XFADE_HOP_LE


def chon_chuyen_canh(so_chuyen):
    """Mỗi video chọn ngẫu nhiên một 'bộ' 6-10 kiểu, không lặp liên tiếp, không trùng đầu với video trước."""
    pool = ds_chuyen_canh_hop_le()
    bo = random.sample(pool, k=min(len(pool), random.randint(6, 10)))
    seq, truoc = [], _TRANG_THAI_HIEU_UNG["chuyen_cuoi"]
    for _ in range(so_chuyen):
        ung_vien = [t for t in bo if t != truoc] or bo
        t = random.choice(ung_vien)
        seq.append(t)
        truoc = t
    _TRANG_THAI_HIEU_UNG["chuyen_cuoi"] = truoc
    return seq


def bieu_thuc_chuyen_dong(kieu, so_khung):
    """Trả về (z, x, y) cho filter zoompan (hiệu ứng Ken Burns)."""
    z_max = 1 + HE_SO_ZOOM
    n = max(so_khung - 1, 1)
    cx = "iw/2-(iw/zoom/2)"
    cy = "ih/2-(ih/zoom/2)"
    if kieu == "in":
        return f"min(1+{HE_SO_ZOOM}*on/{n},{z_max})", cx, cy
    if kieu == "out":
        return f"max({z_max}-{HE_SO_ZOOM}*on/{n},1)", cx, cy
    if kieu == "trai_phai":
        return f"{z_max}", f"(iw-iw/zoom)*on/{n}", cy
    if kieu == "phai_trai":
        return f"{z_max}", f"(iw-iw/zoom)*(1-on/{n})", cy
    if kieu == "tren_duoi":
        return f"{z_max}", cx, f"(ih-ih/zoom)*on/{n}"
    return f"{z_max}", cx, f"(ih-ih/zoom)*(1-on/{n})"  # duoi_tren


def ghep_anh_vua_khung(raw_path, out_png, W, H):
    """Giữ NGUYÊN toàn bộ ảnh (không cắt), phần thừa được lấp bằng nền mờ chính ảnh đó."""
    w8, h8 = max(2, W // 8), max(2, H // 8)
    fc = (
        f"[0:v]split=2[a][b];"
        f"[a]scale={w8}:{h8}:force_original_aspect_ratio=increase,crop={w8}:{h8},"
        f"gblur=sigma=4,scale={W}:{H},eq=brightness=-0.08[bg];"
        f"[b]scale={W}:{H}:force_original_aspect_ratio=decrease[fg];"
        f"[bg][fg]overlay=(W-w)/2:(H-h)/2,setsar=1,format=rgb24[o]"
    )
    chay_ffmpeg(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", raw_path,
                 "-filter_complex", fc, "-map", "[o]", "-frames:v", "1", out_png])


def chuan_bi_video_nen_tu_anh(links, audio_duration, tag, P):
    W, H = P.get("nen_w", P["W"]), P.get("nen_h", P["H"])
    file_tam = []
    try:
        # 1) Tải ảnh và dựng khung vừa kích thước (không cắt ảnh), ảnh lỗi thì bỏ qua
        comps = []
        for idx, url in enumerate(links):
            raw = f"img_{tag}_{idx}"
            comp = f"comp_{tag}_{idx}.png"
            file_tam += [raw, comp]
            try:
                tai_anh_ve(url, raw)
                ghep_anh_vua_khung(raw, comp, W, H)
                comps.append(comp)
            except Exception as e:
                print(f"[ANH] Bỏ qua ảnh {idx + 1} ({url[:70]}): {type(e).__name__}: {str(e)[:120]}", flush=True)
        if not comps:
            raise RuntimeError("Không xử lý được ảnh nào trong danh sách link")

        # 2) Tính số cảnh và độ dài mỗi cảnh (ảnh lặp lại với chuyển động khác nhau)
        n = len(comps)
        t = round(random.uniform(CHUYEN_CANH_MIN, CHUYEN_CANH_MAX), 2)
        shot_giay = random.uniform(SHOT_GIAY_MIN, SHOT_GIAY_MAX)
        m = min(max(n, math.ceil(audio_duration / shot_giay)), MAX_SHOT)
        d = (audio_duration + 1.0 + (m - 1) * t) / m
        d = max(d, 2 * t + 0.6, 2.5)
        so_khung = int(round(d * FPS))
        d = so_khung / FPS

        # 3) Mỗi cảnh = 1 ảnh + 1 kiểu chuyển động (phóng to/thu nhỏ/lướt trái-phải-trên-dưới)
        clips, truoc = [], None
        for k in range(m):
            comp = comps[k % n]
            kieu = random.choice([x for x in KIEU_CHUYEN_DONG if x != truoc])
            truoc = kieu
            z, x, y = bieu_thuc_chuyen_dong(kieu, so_khung)
            clip = f"shot_{tag}_{k}.mp4"
            clips.append(clip)
            file_tam.append(clip)
            vf = (f"scale={int(W * 1.5)}:{int(H * 1.5)},"
                  f"zoompan=z='{z}':x='{x}':y='{y}':d={so_khung}:s={W}x{H}:fps={FPS},"
                  f"setsar=1,format=yuv420p")
            chay_ffmpeg(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", comp,
                         "-vf", vf, "-frames:v", str(so_khung),
                         "-c:v", "libx264", "-preset", "veryfast", "-crf", "23", clip])

        # 4) Nối các cảnh bằng chuyển cảnh ngẫu nhiên
        out = f"bgimg_{tag}.mp4"
        if m == 1:
            os.replace(clips[0], out)
            file_tam.remove(clips[0])
            return out

        seq = chon_chuyen_canh(m - 1)
        parts, prev = [], "[0:v]"
        for k in range(1, m):
            nhan = "[bgv]" if k == m - 1 else f"[x{k}]"
            parts.append(
                f"{prev}[{k}:v]xfade=transition={seq[k - 1]}:duration={t}:offset={k * (d - t):.3f}{nhan}"
            )
            prev = nhan
        cmd = ["ffmpeg", "-y", "-hide_banner", "-loglevel", "error"]
        for c in clips:
            cmd += ["-i", c]
        cmd += ["-filter_complex", ";".join(parts), "-map", "[bgv]", "-an",
                "-c:v", "libx264", "-preset", "veryfast", "-crf", "23", "-pix_fmt", "yuv420p", out]
        chay_ffmpeg(cmd)
        dem = {}
        for s in seq:
            dem[s] = dem.get(s, 0) + 1
        print(f"[HIEU UNG] {m} cảnh x {d:.1f}s, chuyển cảnh {t}s, kiểu dùng: {dem}", flush=True)
        return out
    finally:
        for p in file_tam:
            if p and os.path.exists(p):
                os.remove(p)


# ================== GHÉP CUỐI ==================

def filter_video_ngan(P, ass_esc):
    W, H = P["W"], P["H"]
    k = W / 1080

    def px(v):
        return max(1, round(v * k))

    sigma = max(1.0, 25 * k / 2)
    return (
        f"[0:v]split=2[bgsrc][fgsrc];"
        f"[bgsrc]scale={W // 2}:{H // 2}:force_original_aspect_ratio=increase,"
        f"crop={W // 2}:{H // 2},gblur=sigma={sigma:.1f},scale={W}:{H},eq=brightness=-0.05[bgblur];"
        f"[fgsrc]scale={W}:-2[fgvid];"
        f"[bgblur][fgvid]overlay=(W-w)/2:(H-h)/2[bg];"
        f"[1:v]scale={px(340)}:{px(140)}[logo];"
        f"[bg][logo]overlay=(W-w)/2:{px(55)}[bg2];"
        f"[bg2]drawtext=fontfile={FONT_BOLD}:text='{P['ten_text']}':"
        f"fontsize={px(40)}:fontcolor={P['mau_ten']}:"
        f"borderw={px(3)}:bordercolor=black@0.9:"
        f"box=1:boxcolor=black@0.45:boxborderw={px(18)}:"
        f"x=(w-text_w)/2:y={px(205)}[bg2b];"
        f"[bg2b]drawtext=fontfile={FONT_BOLD}:text='{P['lien_he_text']}':"
        f"fontsize={px(48)}:fontcolor={P['mau_sdt']}:"
        f"borderw={px(3)}:bordercolor=black@0.9:"
        f"box=1:boxcolor=black@0.5:boxborderw={px(20)}:"
        f"x=(w-text_w)/2:y={px(300)}[bg3];"
        f"[bg3]subtitles=filename='{ass_esc}':fontsdir=/usr/share/fonts/truetype/dejavu[outv]"
    )


def filter_video_dai(P, ass_esc):
    return (
        f"[0:v]scale={P['W']}:{P['H']}[bg];"
        f"[1:v]scale={P['logo_w']}:-1[logo];"
        f"[bg][logo]overlay=W-w-20:{P['logo_y']}[bg2];"
        f"[bg2]drawtext=fontfile={FONT_BOLD}:text='{TEXT_LIEN_HE}':"
        f"fontsize={P['contact_size']}:fontcolor={MAU_CHU}:"
        f"borderw=2:bordercolor=black@0.8:box=1:boxcolor=black@0.4:boxborderw=14:"
        f"x=(w-text_w)/2:y={P['contact_y']}[bg3];"
        f"[bg3]subtitles=filename='{ass_esc}':fontsdir=/usr/share/fonts/truetype/dejavu[outv]"
    )


def ghep_video(audio_path, background_video, ass_path, out_path, P):
    ass_abs = os.path.abspath(ass_path)
    if not os.path.exists(ass_abs):
        raise FileNotFoundError(f"Không tìm thấy file phụ đề: {ass_abs}")
    ass_esc = ass_abs.replace(":", r"\:")

    fc = filter_video_ngan(P, ass_esc) if P.get("khung_mo") else filter_video_dai(P, ass_esc)
    chay_ffmpeg([
        "ffmpeg", "-y", "-hide_banner", "-loglevel", "error",
        "-i", background_video, "-i", LOGO_PATH, "-i", audio_path,
        "-filter_complex", fc,
        "-map", "[outv]", "-map", "2:a",
        "-c:v", "libx264", "-preset", "veryfast", "-crf", "23",
        "-pix_fmt", "yuv420p", "-c:a", "aac",
        "-shortest", out_path,
    ])


def upload_to_drive(file_path, file_name, parent_id):
    media = MediaFileUpload(file_path, resumable=True)
    f = drive_upload.files().create(
        body={"name": file_name, "parents": [parent_id]}, media_body=media, fields="id"
    ).execute()
    drive_upload.permissions().create(
        fileId=f["id"], body={"type": "anyone", "role": "reader"}
    ).execute()
    return f"https://drive.google.com/file/d/{f['id']}/view"


# ================== HÀNG ĐỢI XEN KẼ ==================

def lay_viec_cho():
    """Đọc lam_video_dai và lam_video_ngan, xếp xen kẽ 1 dài - 1 ngắn."""
    theo_tab = []
    for ten_tab in PROFILES:
        try:
            ws = retry(spreadsheet.worksheet, ten_tab)
        except gspread.WorksheetNotFound:
            print(f"Không có tab {ten_tab}, bỏ qua.", flush=True)
            continue
        rows = retry(ws.get_all_values)
        ds = []
        for i, row in enumerate(rows[1:], start=2):
            row = row + [""] * (6 - len(row))
            tieu_de = row[0].strip()
            text = row[1].strip()
            loai_den = row[2].strip()
            links = tach_link_anh(row[3])
            status = row[5].strip().lower()
            if status in DONE:
                continue
            if not (tieu_de and text and loai_den and links):
                continue
            if text == "Đang tạo nội dung...":
                continue
            ds.append(dict(
                tab=ten_tab, ws=ws, i=i,
                tieu_de=tieu_de, text=text, loai_den=loai_den, links=links
            ))
        print(f"[{ten_tab}] {len(ds)} video cần làm.", flush=True)
        theo_tab.append(ds)

    xen_ke = []
    for nhom in zip_longest(*theo_tab):
        xen_ke.extend(v for v in nhom if v is not None)
    return xen_ke


# ================== LÀM 1 VIDEO ==================

cache_thu_muc_out = {}


def lay_thu_muc_dich_drive(ten_tab, loai_den):
    """Tạo/lấy thư mục theo cấu trúc: [DRIVE_OUTPUT_FOLDER_ID] / [Video Dai hoặc Video Ngan] / [LoaiDen]."""
    P = PROFILES[ten_tab]
    loai_folder = P["thu_muc_goc"]
    if loai_folder not in cache_thu_muc_out:
        cache_thu_muc_out[loai_folder] = retry(
            tim_hoac_tao_thu_muc, loai_folder, DRIVE_OUTPUT_FOLDER_ID, drive_upload
        )

    ten_loai_den = loai_den.replace("\\", "-").replace("/", "-").strip() or "Khac"
    khoa = (loai_folder, ten_loai_den)
    if khoa not in cache_thu_muc_out:
        cache_thu_muc_out[khoa] = retry(
            tim_hoac_tao_thu_muc, ten_loai_den, cache_thu_muc_out[loai_folder], drive_upload
        )
    return cache_thu_muc_out[khoa]


def lam_1_video(job):
    ten_tab, ws, i = job["tab"], job["ws"], job["i"]
    tieu_de, text, loai_den, links = job["tieu_de"], job["text"], job["loai_den"], job["links"]
    P = PROFILES[ten_tab]

    audio_path = f"audio_{ten_tab}_{i}.mp3"
    ass_path = f"caption_{ten_tab}_{i}.ass"
    out_path = f"output_{ten_tab}_{i}.mp4"
    bg = None
    try:
        words = tao_giong_doc(text, audio_path, P["voices"], P.get("vieneu_voice", "Hải Đăng"))
        viet_phu_de(words, ass_path, P)
        bg = chuan_bi_video_nen_tu_anh(links, get_duration(audio_path), f"{ten_tab}_{i}", P)
        ghep_video(audio_path, bg, ass_path, out_path, P)

        parent_dich = lay_thu_muc_dich_drive(ten_tab, loai_den)
        link = retry(upload_to_drive, out_path, slugify_vi(tieu_de) + ".mp4", parent_dich)

        retry(ws.batch_update, [
            {"range": f"E{i}:F{i}", "values": [[link, "Done"]]},
            {"range": f"I{i}", "values": [[""]]},
        ], value_input_option="USER_ENTERED")
        print(f"[{ten_tab}] Dòng {i}: xong ({P['thu_muc_goc']}/{loai_den}) -> {link}", flush=True)
    except TTSBlocked:
        raise
    except Exception as e:
        print(f"[{ten_tab}] Dòng {i}: LỖI {type(e).__name__}: {e}", flush=True)
        ghi_log(ws, i, f"❌ {type(e).__name__}: {str(e)[:140]}")
    finally:
        for p in (audio_path, ass_path, bg, out_path):
            if p and os.path.exists(p):
                os.remove(p)


# ================== MAIN ==================

def main():
    bat_dau = time.time()
    try:
        ip = urllib.request.urlopen("https://api.ipify.org", timeout=10).read().decode()
        print(f"IP máy chủ GitHub lần này: {ip} | edge-tts {getattr(edge_tts, '__version__', '?')}", flush=True)
    except Exception:
        pass

    khoi_tao()
    viec = lay_viec_cho()
    print(f"Tổng cộng {len(viec)} video, làm xen kẽ giữa các tab.", flush=True)

    da_lam, bi_chan_tts = 0, False
    for job in viec:
        if (time.time() - bat_dau) / 60 > GIOI_HAN_PHUT:
            print("Hết thời gian của lần chạy này, phần còn lại để lần sau.", flush=True)
            break
        if MAX_VIDEO and da_lam >= MAX_VIDEO:
            print(f"Đã đủ {MAX_VIDEO} video, dừng.", flush=True)
            break
        try:
            lam_1_video(job)
        except TTSBlocked as e:
            print(f"Không có giọng đọc nào dùng được, dừng sớm: {e}", flush=True)
            bi_chan_tts = True
            break
        except Exception as e:
            print(f"[{job['tab']}] Dòng {job['i']}: LỖI NGOÀI DỰ KIẾN {type(e).__name__}: {e}", flush=True)
        da_lam += 1

    print(f"Đã xử lý {da_lam} video.", flush=True)
    if bi_chan_tts:
        sys.exit(1)


if __name__ == "__main__":
    main()
