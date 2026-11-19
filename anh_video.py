# -*- coding: utf-8 -*-
"""Dựng video nền từ 1-5 ảnh (link web). Mọi tham số đọc từ tab CONFIG của Google Sheet.
Dùng cùng main.py: main.py lo giọng đọc, phụ đề, logo, upload; file này chỉ tạo "video nền" từ ảnh."""
import io
import os
import re
import random
import shutil
import subprocess
import time
import urllib.parse
import urllib.request
from datetime import datetime

import gspread
from PIL import Image, ImageEnhance, ImageFilter, ImageOps

Image.MAX_IMAGE_PIXELS = 60_000_000
TAB_CONFIG = "CONFIG"
TAB_NHAT_KY = "nhat_ky_hieu_ung"

MAC_DINH = {
    "KHUNG_ANH": "ngang",
    "GIONG_ANH": "nu",
    "MAX_ANH": "5",
    "SO_TU_MOI_ANH": "40",
    "FPS": "30",
    "GIAY_CANH_MIN": "3.5",
    "GIAY_CANH_MAX": "6",
    "GIAY_CHUYEN_MIN": "0.6",
    "GIAY_CHUYEN_MAX": "1.1",
    "ZOOM_MAX": "1.25",
    "DS_CHUYEN_CANH": "fade,fadeblack,fadegrays,dissolve,wipeleft,wiperight,wipeup,wipedown,slideleft,slideright,"
                      "slideup,slidedown,smoothleft,smoothright,circleopen,circleclose,radial,pixelize,diagtl,diagbr,"
                      "horzopen,vertopen,hlslice,vuslice,squeezeh,zoomin",
    "DS_CHUYEN_DONG": "zoom_in,zoom_out,pan_trai,pan_phai,pan_len,pan_xuong,zoom_in_goc,zoom_out_goc",
    "BIEN_THE_MAU": "TRUE",
    "NANG_CAP_ANH_WOO": "TRUE",
    "SO_VIDEO_TRANH_TRUNG": "20",
    "MAX_VIDEO": "0",
    "GIOI_HAN_PHUT": "300",
}
XFADE_HOP_LE = set(MAC_DINH["DS_CHUYEN_CANH"].split(",")) | {"hblur", "fadewhite", "rectcrop", "circlecrop", "distance"}
DONG_HOP_LE = {"zoom_in", "zoom_out", "pan_trai", "pan_phai", "pan_len", "pan_xuong", "zoom_in_goc", "zoom_out_goc"}

CFG = dict(MAC_DINH)


# ---------------- CONFIG ----------------
def doc_config(spreadsheet):
    global CFG
    CFG = dict(MAC_DINH)
    try:
        ws = spreadsheet.worksheet(TAB_CONFIG)
        for row in ws.get_all_values()[1:]:
            if len(row) >= 2 and row[0].strip() and row[1].strip() != "":
                CFG[row[0].strip().upper()] = row[1].strip()
        print(f"[CONFIG] đã đọc {len(CFG)} tham số từ tab {TAB_CONFIG}", flush=True)
    except gspread.WorksheetNotFound:
        print("[CONFIG] không có tab CONFIG, dùng giá trị mặc định.", flush=True)
    return CFG


def _s(k):
    return str(CFG.get(k, MAC_DINH.get(k, ""))).strip()


def cf(k):
    try:
        return float(_s(k).replace(",", "."))
    except ValueError:
        return float(MAC_DINH[k])


def ci(k):
    return int(cf(k))


def cb(k):
    return _s(k).lower() in ("true", "1", "yes", "có", "co", "x")


def cl(k):
    return [x.strip() for x in re.split(r"[,;\n]+", _s(k)) if x.strip()]


def ap_cau_hinh(PROFILES, giong_nu, giong_nam):
    """Tạo profile cho tab lam_video_anh dựa trên CONFIG: khung ngang hoặc dọc (kiểu Shorts), giọng đọc."""
    khung = _s("KHUNG_ANH").lower()
    goc = PROFILES["lam_video_ngan" if khung in ("doc", "dọc", "ngan", "ngắn", "9:16") else "lam_video_dai"]
    p = dict(goc)
    p.update(kieu="anh", thu_muc="Video Anh")
    gioi = _s("GIONG_ANH").lower()
    if gioi == "nam":
        p["voices"], p["vieneu_voice"] = [giong_nam, giong_nu], "Hải Đăng"
    else:
        p["voices"], p["vieneu_voice"] = [giong_nu, giong_nam], "Trúc Ly"
    p["giong_ngau_nhien"] = gioi in ("ngau_nhien", "ngẫu nhiên", "random")
    PROFILES["lam_video_anh"] = p


def doc_link_anh(cells):
    ds = []
    for c in cells:
        for u in re.split(r"[\s,;|]+", str(c).strip()):
            if u.lower().startswith("http") and u not in ds:
                ds.append(u)
    return ds[:max(1, min(ci("MAX_ANH"), 5))]


# ---------------- TẢI + XỬ LÝ ẢNH ----------------
def tai_anh(url):
    ung_vien = []
    if cb("NANG_CAP_ANH_WOO"):
        goc = re.sub(r"-\d{2,4}x\d{2,4}(?=\.(?:jpe?g|png|webp)(?:\?.*)?$)", "", url, flags=re.I)
        if goc != url:
            ung_vien.append(goc)   # ảnh gốc to, đẹp hơn bản thumbnail
    ung_vien.append(url)
    for u in ung_vien:
        try:
            u2 = urllib.parse.quote(u, safe=":/?&=%#@+~,;")
            req = urllib.request.Request(u2, headers={"User-Agent": "Mozilla/5.0"})
            data = urllib.request.urlopen(req, timeout=30).read()
            im = ImageOps.exif_transpose(Image.open(io.BytesIO(data))).convert("RGB")
            if min(im.size) < 120:
                continue
            im.thumbnail((3000, 3000))
            return im
        except Exception as e:
            print(f"[ANH] không tải được {u[:90]}: {type(e).__name__} {str(e)[:60]}", flush=True)
    return None


def cat_chi_tiet(img, rng):
    w, h = img.size
    f = rng.uniform(0.6, 0.78)
    cw, ch = int(w * f), int(h * f)
    x = int((w - cw) * rng.uniform(0.2, 0.8))
    y = int((h - ch) * rng.uniform(0.2, 0.8))
    return img.crop((x, y, x + cw, y + ch))


def dung_khung(img, CW, CH, rng):
    """Ảnh gần đúng tỉ lệ khung -> phủ kín. Ảnh lệch tỉ lệ (vuông, dọc) -> ảnh sắc nét ở giữa + nền là chính ảnh làm mờ."""
    iw, ih = img.size
    if abs((iw / ih) / (CW / CH) - 1) < 0.15:
        return ImageOps.fit(img, (CW, CH), Image.LANCZOS)
    nen = ImageOps.fit(img, (CW // 4, CH // 4), Image.BILINEAR).filter(ImageFilter.GaussianBlur(rng.uniform(6, 11)))
    nen = nen.resize((CW, CH), Image.BICUBIC)
    nen = ImageEnhance.Brightness(nen).enhance(rng.uniform(0.7, 0.9))
    ty_le = rng.uniform(0.86, 0.94)
    fg = ImageOps.contain(img, (int(CW * ty_le), int(CH * ty_le)), Image.LANCZOS)
    nen.paste(fg, ((CW - fg.width) // 2, (CH - fg.height) // 2))
    return nen


# ---------------- KẾ HOẠCH HIỆU ỨNG ----------------
def lap_ke_hoach(n_anh, muc_tieu, fps, rng):
    ds_chuyen = [t for t in cl("DS_CHUYEN_CANH") if t in XFADE_HOP_LE] or ["fade"]
    ds_dong = [m for m in cl("DS_CHUYEN_DONG") if m in DONG_HOP_LE] or ["zoom_in"]
    bo_chuyen = rng.sample(ds_chuyen, min(len(ds_chuyen), rng.randint(4, 7)))
    bo_dong = rng.sample(ds_dong, min(len(ds_dong), rng.randint(4, 6)))
    dmin, dmax = cf("GIAY_CANH_MIN"), cf("GIAY_CANH_MAX")
    tmin, tmax = cf("GIAY_CHUYEN_MIN"), cf("GIAY_CHUYEN_MAX")
    can = (muc_tieu + (n_anh - 1) * tmax) / n_anh      # lời đọc ngắn -> rút ngắn cảnh để ảnh nào cũng kịp lên hình
    if can < dmin:
        dmin = max(2.0, can * 0.9)
        dmax = max(dmin + 0.5, min(dmax, can))
    shots, dung, tong = [], [0] * n_anh, 0.0
    while len(shots) < 80:
        k = len(shots)
        d = max(1, round(rng.uniform(dmin, dmax) * fps)) / fps
        if k < n_anh:
            idx = k
        elif n_anh == 1:
            idx = 0
        else:
            ung = [j for j in range(n_anh) if j != shots[-1]["anh"]]
            it = min(dung[j] for j in ung)
            idx = rng.choice([j for j in ung if dung[j] == it])
        prev_m = shots[-1]["dong"] if shots else None
        dong = rng.choice([m for m in bo_dong if m != prev_m] or bo_dong)
        chuyen, td = None, 0.0
        if shots:
            chuyen = rng.choice([t for t in bo_chuyen if t != shots[-1]["chuyen"]] or bo_chuyen)
            td = round(min(rng.uniform(tmin, tmax), d / 2.2, shots[-1]["dur"] / 2.2), 2)
        shots.append(dict(anh=idx, lan=dung[idx], dong=dong, dur=d, chuyen=chuyen, td=td))
        dung[idx] += 1
        tong += d - td
        if tong >= muc_tieu and len(shots) >= n_anh:
            break
    return shots


def tokens(ke_hoach):
    return [f"{s['chuyen'] or '-'}/{s['dong']}" for s in ke_hoach]


def tuong_dong(a, b):
    if not a or not b:
        return 0.0
    n = min(len(a), len(b))
    vi_tri = sum(x == y for x, y in zip(a[:n], b[:n])) / n
    sa, sb = set(a), set(b)
    return 0.5 * vi_tri + 0.5 * len(sa & sb) / len(sa | sb)


def doc_lich_su(spreadsheet):
    k = ci("SO_VIDEO_TRANH_TRUNG")
    if not spreadsheet or k <= 0:
        return []
    try:
        vals = spreadsheet.worksheet(TAB_NHAT_KY).col_values(4)[1:]
    except Exception:
        return []
    return [v.split(";")[0].split("|") for v in vals[-k:] if v]


def ghi_nhat_ky(spreadsheet, tieu_de, so_anh, chu_ky):
    try:
        try:
            ws = spreadsheet.worksheet(TAB_NHAT_KY)
        except gspread.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title=TAB_NHAT_KY, rows=1000, cols=4)
            ws.append_row(["Thời gian", "Tiêu đề", "Số ảnh", "Chữ ký hiệu ứng"])
        ws.append_row([datetime.now().strftime("%d/%m/%Y %H:%M:%S"), tieu_de, so_anh, chu_ky])
    except Exception as e:
        print(f"[ANH] không ghi được nhật ký hiệu ứng: {e}", flush=True)


# ---------------- RENDER ----------------
def _ff(cmd):
    r = subprocess.run(cmd, capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"ffmpeg lỗi: {r.stderr[-300:]}")


def bieu_thuc(dong, n, zm, rng):
    p = f"on/{n}"
    cx, cy = "iw/2-iw/zoom/2", "ih/2-ih/zoom/2"
    zc = f"{rng.uniform(1.12, max(zm, 1.13)):.3f}"
    dz = f"{zm - 1:.3f}"
    fx, fy = rng.choice([0.05, 0.95]), rng.choice([0.05, 0.95])
    gx, gy = f"(iw-iw/zoom)*{fx}", f"(ih-ih/zoom)*{fy}"
    return {
        "zoom_in": (f"1+{dz}*{p}", cx, cy),
        "zoom_out": (f"{zm:.3f}-{dz}*{p}", cx, cy),
        "pan_trai": (zc, f"(iw-iw/zoom)*(1-{p})", cy),
        "pan_phai": (zc, f"(iw-iw/zoom)*{p}", cy),
        "pan_len": (zc, cx, f"(ih-ih/zoom)*(1-{p})"),
        "pan_xuong": (zc, cx, f"(ih-ih/zoom)*{p}"),
        "zoom_in_goc": (f"1+{dz}*{p}", gx, gy),
        "zoom_out_goc": (f"{zm:.3f}-{dz}*{p}", gx, gy),
    }[dong]


def render_canh(frame, out, dong, dur, W, H, fps, zm, mau, rng):
    n = max(1, int(round(dur * fps)))
    z, x, y = bieu_thuc(dong, n, zm, rng)
    vf = (f"zoompan=z='{z}':x='{x}':y='{y}':d={n}:s={W}x{H}:fps={fps},"
          f"eq=brightness={mau[0]:.3f}:contrast={mau[1]:.3f}:saturation={mau[2]:.3f},setsar=1,format=yuv420p")
    _ff(["ffmpeg", "-y", "-hide_banner", "-loglevel", "error", "-i", frame, "-vf", vf,
         "-frames:v", str(n), "-c:v", "libx264", "-preset", "veryfast", "-crf", "22", out])


def noi_canh(clips, shots, out):
    if len(clips) == 1:
        shutil.copyfile(clips[0], out)
        return
    cmd = ["ffmpeg", "-y", "-hide_banner", "-loglevel", "error"]
    for c in clips:
        cmd += ["-i", c]
    parts, last, L = [], "[0:v]", shots[0]["dur"]
    for k in range(1, len(clips)):
        s = shots[k]
        off = L - s["td"]
        parts.append(f"{last}[{k}:v]xfade=transition={s['chuyen']}:duration={s['td']}:offset={off:.3f}[x{k}]")
        last, L = f"[x{k}]", off + s["dur"]
    cmd += ["-filter_complex", ";".join(parts), "-map", last, "-an", "-c:v", "libx264",
            "-preset", "veryfast", "-crf", "22", "-pix_fmt", "yuv420p", out]
    _ff(cmd)


def dung_video_anh(links, audio_giay, tag, P, tieu_de, spreadsheet=None):
    """Trả về (đường dẫn video nền, chữ ký hiệu ứng). Video nền dài hơn giọng đọc ~1.5s."""
    W, H = P.get("nen_w", P["W"]), P.get("nen_h", P["H"])
    fps = ci("FPS") or 30
    rng = random.Random(f"{tieu_de}|{time.time_ns()}")
    imgs = [im for im in (tai_anh(u) for u in links) if im]
    if not imgs:
        raise RuntimeError("Không tải được ảnh nào từ các link trong Sheet")
    print(f"[ANH] dùng {len(imgs)}/{len(links)} ảnh", flush=True)

    lich_su, muc_tieu, tot = doc_lich_su(spreadsheet), audio_giay + 1.5, None
    for _ in range(8):   # thử 8 kế hoạch, chọn cái khác nhất với các video gần đây
        kh = lap_ke_hoach(len(imgs), muc_tieu, fps, rng)
        diem = max([tuong_dong(tokens(kh), h) for h in lich_su] or [0.0])
        if tot is None or diem < tot[0]:
            tot = (diem, kh)
        if diem < 0.3:
            break
    shots = tot[1]
    base = (rng.uniform(-0.03, 0.03), rng.uniform(0.97, 1.06), rng.uniform(0.95, 1.12)) if cb("BIEN_THE_MAU") else (0, 1, 1)
    zm = max(1.1, cf("ZOOM_MAX"))
    out = f"bgconcat_{tag}.mp4"
    tam = []
    try:
        for k, s in enumerate(shots):
            im = imgs[s["anh"]]
            if s["lan"] > 0:
                im = cat_chi_tiet(im, rng)    # lần xuất hiện thứ 2 trở đi: cắt cận chi tiết khác
            fr, cl_ = f"frame_{tag}_{k}.jpg", f"shot_{tag}_{k}.mp4"
            tam += [fr, cl_]
            dung_khung(im, W * 2, H * 2, rng).save(fr, quality=92)
            mau = tuple(b + rng.uniform(-d, d) for b, d in zip(base, (0.01, 0.01, 0.02)))
            render_canh(fr, cl_, s["dong"], s["dur"], W, H, fps, zm, mau, rng)
            os.remove(fr)
        noi_canh([f"shot_{tag}_{k}.mp4" for k in range(len(shots))], shots, out)
    finally:
        for p in tam:
            if os.path.exists(p):
                os.remove(p)
    chu_ky = "|".join(tokens(shots)) + f";mau={base[0]:+.2f}/{base[1]:.2f}/{base[2]:.2f}"
    print(f"[ANH] {len(shots)} cảnh, độ giống video trước: {tot[0]:.2f}", flush=True)
    return out, chu_ky
