#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
TẢI VIDEO HÀNG LOẠT TỪ FILE EXCEL
---------------------------------
- Đọc file Excel, tìm cột "Link video" và cột "Đặt tên file tải về".
- Tải từng link bằng yt-dlp (YouTube, TikTok, Facebook, Instagram, Pinterest,
  X/Twitter, báo điện tử... khoảng 1.800 trang).
- Lưu file với đúng tên trong Excel, ví dụ 2.1.mp4, 2.2.mp4...
- File đã tải rồi sẽ được bỏ qua, nên chạy lại chỉ tải tiếp các link còn lỗi.
- Ghi báo cáo kết quả ra file CSV trong thư mục video.

Cách dùng:
  python tai_video.py                      -> mở hộp thoại chọn file Excel
  python tai_video.py danh_sach.xlsx       -> chạy với file chỉ định
  python tai_video.py danh_sach.xlsx D:\Video   -> chỉ định thêm thư mục lưu
"""

import csv
import datetime
import os
import re
import shutil
import sys
import tempfile
import unicodedata
from pathlib import Path

# ===================== CẤU HÌNH (có thể sửa) =====================
# Đăng nhập cho Facebook/Instagram (video riêng tư hoặc bị chặn khi chưa đăng nhập):
#   Cách 1: đặt file cookies.txt cạnh file này (khuyên dùng, xem HUONG_DAN.txt).
#   Cách 2: lấy cookie trực tiếp từ trình duyệt: "firefox", "edge", "chrome"...
#           Để None nếu không dùng.
COOKIES_FROM_BROWSER = None

# Chất lượng tối đa (chiều cao, pixel). None = cao nhất có thể.
MAX_HEIGHT = None

# Tự chuyển video về MP4 mã H.264 (Premiere và Windows đọc được ngay). False = giữ nguyên file tải về.
CHUYEN_H264 = True

# Trang bắt đăng nhập mới xem được video: nhờ dịch vụ trung gian lấy hộ, chỉ gửi địa chỉ/mã bài đăng đó đi.
#   TikTok -> tikwm.com ; X/Twitter -> FxTwitter (api.fxtwitter.com)
# False = không bao giờ gửi link cho bên thứ ba; các link này sẽ báo lỗi "cần đăng nhập".
TIKTOK_TRUNG_GIAN = True

# Từ khóa nhận diện tiêu đề cột (không phân biệt hoa thường/dấu)
LINK_HEADER_KEYS = ["link video", "link", "url"]
NAME_HEADER_KEYS = ["dat ten file", "ten file", "file name", "filename"]
# Cột dự phòng nếu không tìm thấy tiêu đề (C = link, D = tên file)
FALLBACK_LINK_COL = 3
FALLBACK_NAME_COL = 4
# ================================================================

PHIEN_BAN = "16"
SCRIPT_DIR = Path(__file__).resolve().parent
VIDEO_EXTS = {".mp4", ".mkv", ".webm", ".mov", ".m4v", ".flv", ".avi"}


def bo_dau(text: str) -> str:
    text = unicodedata.normalize("NFD", str(text))
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return text.replace("đ", "d").replace("Đ", "D").lower().strip()


def ten_hop_le(name: str) -> str:
    """Bỏ các ký tự Windows không cho phép trong tên file."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name).strip().rstrip(". ")
    return name[:150]


def doc_ten_file(cell) -> str:
    """Đọc ô tên file, giữ nguyên số thập phân như Excel hiển thị (2.1, 8.10...)."""
    v = cell.value
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        fmt = cell.number_format or ""
        m = re.search(r"0\.(0+)", fmt)
        if m:
            return f"{v:.{len(m.group(1))}f}"
        if v.is_integer():
            return str(int(v))
        return repr(v)
    return str(v).strip()


def doc_link(cell) -> str:
    if cell.hyperlink is not None and cell.hyperlink.target:
        return cell.hyperlink.target.strip()
    v = cell.value
    if v is None:
        return ""
    s = str(v).strip()
    # Ô có công thức =HYPERLINK("url", "text")
    m = re.match(r'^=HYPERLINK\(\s*"([^"]+)"', s, re.I)
    if m:
        return m.group(1)
    return s


def tim_cot(ws):
    """Tìm hàng tiêu đề và vị trí 2 cột trong 30 hàng đầu."""
    for row in ws.iter_rows(min_row=1, max_row=min(30, ws.max_row)):
        link_col = name_col = None
        for c in row:
            if c.value is None:
                continue
            t = bo_dau(c.value)
            if name_col is None and any(k in t for k in NAME_HEADER_KEYS):
                name_col = c.column
            elif link_col is None and any(k == t or t.startswith(k) for k in LINK_HEADER_KEYS):
                link_col = c.column
        if link_col and name_col:
            return row[0].row, link_col, name_col
    return None


def doc_sheet(ws, header_row, link_col, name_col):
    items = []
    for r in range(header_row + 1, ws.max_row + 1):
        link = doc_link(ws.cell(r, link_col))
        name = doc_ten_file(ws.cell(r, name_col))
        if not link or not link.lower().startswith(("http://", "https://")):
            continue
        if not name:
            print(f"! [{ws.title}] Hàng {r}: có link nhưng thiếu tên file, sẽ đặt tên 'hang_{r}'.")
            name = f"hang_{r}"
        items.append({"row": r, "link": link, "name": ten_hop_le(name)})
    # Tên trùng trong cùng sheet -> thêm hậu tố để không ghi đè
    seen = {}
    for it in items:
        n = it["name"]
        if n in seen:
            seen[n] += 1
            it["name"] = f"{n}_{seen[n]}"
            print(f"! [{ws.title}] Tên '{n}' bị trùng (hàng {it['row']}), đổi thành '{it['name']}'.")
        else:
            seen[n] = 1
    return items


def doc_excel(path: Path):
    """Trả về danh sách (tên sheet, các link) cho MỌI sheet có cột link + tên file."""
    import openpyxl

    # Không dùng read_only để đọc được hyperlink
    wb = openpyxl.load_workbook(path, data_only=False)
    ket_qua = []
    for ws in wb.worksheets:
        if ws.sheet_state != "visible":
            continue
        r = tim_cot(ws)
        if r:
            print(f"Sheet '{ws.title}': cột link {r[1]}, cột tên file {r[2]}")
            ket_qua.append((ws.title, doc_sheet(ws, *r)))
    if not ket_qua:
        ws = wb.active
        print(f"! Không tìm thấy tiêu đề cột, dùng mặc định cột C (link) và D (tên file) ở sheet '{ws.title}'.")
        ket_qua.append((ws.title, doc_sheet(ws, 0, FALLBACK_LINK_COL, FALLBACK_NAME_COL)))
    return [(t, it) for t, it in ket_qua if it]


def chuan_bi_ffmpeg():
    """Tìm ffmpeg (để ghép hình + tiếng chất lượng cao). Trả về thư mục chứa ffmpeg hoặc None."""
    exe = "ffmpeg.exe" if os.name == "nt" else "ffmpeg"
    local = SCRIPT_DIR / "_ffmpeg" / exe
    if local.exists():
        return str(local.parent)
    found = shutil.which("ffmpeg")
    if found:
        return str(Path(found).parent)
    try:
        import imageio_ffmpeg

        src = imageio_ffmpeg.get_ffmpeg_exe()
        local.parent.mkdir(exist_ok=True)
        shutil.copy2(src, local)
        if os.name != "nt":
            os.chmod(local, 0o755)
        return str(local.parent)
    except Exception:
        return None


def da_tai(outdir: Path, name: str):
    for p in outdir.glob(glob_escape(name) + ".*"):
        if p.stem == name and p.suffix.lower() in VIDEO_EXTS and p.stat().st_size > 0:
            return p
    return None


def don_rac(outdir: Path, name: str):
    """Xóa file tạm của lần tải hỏng trước, tránh tải tiếp nhầm từ nguồn khác."""
    pat = re.compile(re.escape(name) + r"(\.f[\w-]+)?\.\w+(\.part|\.ytdl|\.temp)?(-Frag\d+)?$")
    for p in outdir.glob(glob_escape(name) + ".*"):
        if p.stem == name and p.suffix.lower() in VIDEO_EXTS:
            continue
        if pat.fullmatch(p.name) and (p.name.endswith((".part", ".ytdl", ".temp", ".unknown_video"))
                                      or ".f" in p.name[len(name):] or "Frag" in p.name):
            try:
                p.unlink()
            except Exception:
                pass


def glob_escape(s: str) -> str:
    return re.sub(r"([\[\]*?])", r"[\1]", s)


PROGRESS_HOOK = None   # giao diện gán hàm này để nhận tiến độ tải
COOKIE_GOP = None      # file gộp từ cookies.txt + thư mục cookies\*.txt
COOKIE_AUTO = None     # file gộp + cookie tool tự lấy bằng trình duyệt ẩn
_AUTO_COOKIES = {}
TRANG_THAI_TD = "chưa kiểm tra"   # trạng thái dò bằng trình duyệt
TAM_DIR = SCRIPT_DIR / "_tam"

# ----- Chạy nhiều luồng (giao diện tải vài link cùng lúc) -----
import threading  # noqa: E402

_KHOA_COOKIE = threading.Lock()   # ghi/chép file cookie
_KHOA_DUOI = threading.Lock()     # bật/tắt chế độ "cho phép đuôi file lạ" của yt-dlp
_DEM_DUOI = [0, None]             # [số luồng đang cần chế độ đó, hàm gốc của yt-dlp]


def nen_tang(url: str) -> str:
    """Tên nền tảng của link. Giao diện chỉ tải MỘT link mỗi nền tảng tại một thời điểm,
    vì các trang chặn theo địa chỉ mạng khi thấy nhiều yêu cầu dồn dập."""
    from urllib.parse import urlsplit

    try:
        may = (urlsplit(url).hostname or "").lower()
    except Exception:
        may = ""
    for ten, mau in (("douyin", r"douyin\.com$|iesdouyin\.com$"), ("tiktok", r"tiktok\.com$"),
                     ("meta", r"facebook\.com$|fb\.watch$|fb\.com$|instagram\.com$|threads\.(net|com)$"),
                     ("youtube", r"youtube\.com$|youtu\.be$|youtube-nocookie\.com$"),
                     ("x", r"(^|\.)x\.com$|twitter\.com$|t\.co$"), ("reddit", r"reddit\.com$|redd\.it$"),
                     ("pinterest", r"pinterest\.[\w.]+$|pin\.it$"), ("linkedin", r"linkedin\.com$")):
        if re.search(mau, may):
            return ten
    phan = may.split(".")
    # tên miền kiểu dantri.com.vn, bbc.co.uk: lấy 3 phần cuối để không gộp mọi trang .com.vn làm một
    n = 3 if len(phan) >= 3 and len(phan[-1]) == 2 and phan[-2] in ("com", "co", "net", "org", "gov", "edu") else 2
    return ".".join(phan[-n:]) or "khac"


# ----- Proxy HTTP (tùy chọn, đặt trong Cài đặt của giao diện) -----
PROXY_DS = []            # danh sách proxy dạng "http://user:pass@may:cong"
PROXY_CHE_DO = "tat"     # "tat" | "khi_loi" (tải thẳng trước, bị chặn mới đi proxy) | "luon"
_TL = threading.local()  # .proxy = proxy mà luồng này đang dùng cho lần thử hiện tại (None = đi thẳng)
_PROXY_LUOT = [0]
# Lỗi do MẠNG/ĐỊA CHỈ IP (đổi đường đi có thể giúp). Lỗi cần đăng nhập, link sai... thì proxy không giúp được.
_LOI_MANG = re.compile(r"ip address is blocked|not available in your country|geo.?restrict|failed to resolve|getaddrinfo|"
                       r"\bssl\b|eof occurred|\b403\b|forbidden|cloudflare|timed? ?out|connection|\b429\b|too many requests|"
                       r"rate.?limit|douyin không trả dữ liệu|blocked by network", re.I)


def phan_tich_proxy(dong: str):
    """Đọc 1 dòng proxy. Nhận: may:cong:user:pass | may:cong | user:pass@may:cong | http://...
    Trả về "http://user:pass@may:cong" hoặc báo lỗi ValueError."""
    from urllib.parse import quote, urlsplit

    s = dong.strip()
    if not s:
        return None
    if "://" in s:
        sp = urlsplit(s)
        if sp.scheme not in ("http", "https") or not sp.hostname or not sp.port:
            raise ValueError(f"Proxy không hợp lệ (chỉ nhận proxy HTTP): {s[:40]}")
        may, cong, user, mk = sp.hostname, sp.port, sp.username or "", sp.password or ""
        from urllib.parse import unquote
        user, mk = unquote(user), unquote(mk)
    else:
        phan = s.split(":", 3)
        # may:cong:user:pass được xét TRƯỚC, vì mật khẩu có thể chứa ':' và '@'
        if "@" in s and not (len(phan) == 4 and phan[1].isdigit()):
            tk, s2 = s.rsplit("@", 1)
            user, _, mk = tk.partition(":")
            may, _, cong = s2.partition(":")
        else:
            if len(phan) not in (2, 4):
                raise ValueError(f"Proxy phải có dạng may:cong:user:pass (hoặc may:cong): {s[:40]}")
            may, cong = phan[0], phan[1]
            user, mk = (phan[2], phan[3]) if len(phan) == 4 else ("", "")
        if not may or not str(cong).isdigit():
            raise ValueError(f"Proxy phải có dạng may:cong:user:pass (hoặc may:cong): {s[:40]}")
    tk = f"{quote(user, safe='')}:{quote(mk, safe='')}@" if user else ""
    return f"http://{tk}{may}:{int(cong)}"


def che_proxy(p: str) -> str:
    """Tên proxy để ghi nhật ký: không bao giờ in tài khoản/mật khẩu."""
    return p.rsplit("@", 1)[-1].replace("http://", "") if p else ""


def dat_proxy(van_ban: str, che_do: str):
    """Đặt danh sách proxy (mỗi dòng một proxy) và chế độ dùng."""
    global PROXY_DS, PROXY_CHE_DO
    ds = [phan_tich_proxy(d) for d in (van_ban or "").splitlines() if d.strip()]
    if che_do not in ("tat", "khi_loi", "luon"):
        che_do = "tat"
    if che_do != "tat" and not ds:
        raise ValueError("Chưa nhập proxy nào.")
    PROXY_DS, PROXY_CHE_DO = ds, che_do


def proxy_hien_tai():
    return getattr(_TL, "proxy", None)


def _chon_proxy():
    with _KHOA_COOKIE:
        _PROXY_LUOT[0] += 1
        return PROXY_DS[_PROXY_LUOT[0] % len(PROXY_DS)]


def _mo_url(req, timeout):
    """urlopen, đi qua proxy của luồng nếu đang có."""
    import urllib.request

    p = proxy_hien_tai()
    if not p:
        return urllib.request.urlopen(req, timeout=timeout)
    return urllib.request.build_opener(urllib.request.ProxyHandler({"http": p, "https": p})).open(req, timeout=timeout)


def _proxy_cho_trinh_duyet():
    from urllib.parse import urlsplit, unquote

    p = proxy_hien_tai()
    if not p:
        return None
    sp = urlsplit(p)
    d = {"server": f"http://{sp.hostname}:{sp.port}"}
    if sp.username:
        d["username"], d["password"] = unquote(sp.username), unquote(sp.password or "")
    return d


def _ban_sao_cookie(goc: str) -> str:
    """Mỗi luồng dùng một bản sao file cookie riêng: yt-dlp GHI ĐÈ file cookie khi tải xong,
    hai luồng dùng chung một file sẽ làm hỏng file."""
    dst = TAM_DIR / f"ck_{threading.get_ident()}.txt"
    try:
        with _KHOA_COOKIE:
            TAM_DIR.mkdir(exist_ok=True)
            shutil.copyfile(goc, dst)
        return str(dst)
    except Exception:
        return goc


def xoa_ban_sao_cookie():
    try:
        for f in TAM_DIR.glob("ck_*.txt"):
            f.unlink(missing_ok=True)
    except Exception:
        pass


def _doc_dong_cookie(path: Path):
    try:
        txt = path.read_text(encoding="utf-8", errors="ignore")
    except Exception:
        return []
    out = []
    for line in txt.splitlines():
        t = line.strip()
        if not t or (t.startswith("#") and not t.startswith("#HttpOnly_")):
            continue
        if len(line.split("\t")) >= 7:
            out.append(line.rstrip("\r\n"))
    return out


def gop_cookie():
    """Gộp cookies.txt và mọi file trong thư mục cookies\\ thành 1 file tạm.
    (yt-dlp ghi đè file cookie khi chạy xong nên không dùng trực tiếp file gốc.)"""
    global COOKIE_GOP
    files = []
    if (SCRIPT_DIR / "cookies.txt").exists():
        files.append(SCRIPT_DIR / "cookies.txt")
    if (SCRIPT_DIR / "cookies").is_dir():
        files += sorted((SCRIPT_DIR / "cookies").glob("*.txt"))
    lines = []
    for f in files:
        ds = _doc_dong_cookie(f)
        print(f"Cookie: {f.name} ({len(ds)} dòng)" + ("" if ds else "  <- KHÔNG đúng định dạng Netscape cookies.txt"))
        lines += ds
    if not lines:
        COOKIE_GOP = None
        return
    TAM_DIR.mkdir(exist_ok=True)
    dst = TAM_DIR / "cookies_gop.txt"
    dst.write_text("# Netscape HTTP Cookie File\n" + "\n".join(lines) + "\n", encoding="utf-8")
    COOKIE_GOP = str(dst)


def luu_cookie_tu_dong(cookies):
    """Lưu cookie lấy từ trình duyệt ẩn (vd. Douyin cần 'fresh cookies', không cần đăng nhập)."""
    global COOKIE_AUTO
    for c in cookies or []:
        try:
            dom = c["domain"]
            exp = int(c.get("expires") or 0)
            line = "\t".join([dom, "TRUE" if dom.startswith(".") else "FALSE", c.get("path") or "/",
                              "TRUE" if c.get("secure") else "FALSE", str(max(exp, 0)), c["name"], c.get("value", "")])
            _AUTO_COOKIES[(dom, c["name"], c.get("path"))] = line
        except Exception:
            continue
    if not _AUTO_COOKIES:
        return
    with _KHOA_COOKIE:
        TAM_DIR.mkdir(exist_ok=True)
        goc = _doc_dong_cookie(Path(COOKIE_GOP)) if COOKIE_GOP else []
        dst = TAM_DIR / "cookies_auto.txt"
        dst.write_text("# Netscape HTTP Cookie File\n" + "\n".join(list(_AUTO_COOKIES.values()) + goc) + "\n",
                       encoding="utf-8")
        COOKIE_AUTO = str(dst)


def tao_options(outdir: Path, name: str, ffmpeg_dir):
    h = f"[height<={MAX_HEIGHT}]" if MAX_HEIGHT else ""
    if ffmpeg_dir:
        fmt = (f"bv*{h}[ext=mp4]+ba[ext=m4a]/bv*{h}+ba/b{h}[ext=mp4]/b{h}/b")
    else:
        fmt = f"b{h}[ext=mp4]/b{h}/b"
    opts = {
        "outtmpl": str(outdir / f"{name}.%(ext)s"),
        "format": fmt,
        "merge_output_format": "mp4",
        "noplaylist": True,
        "playlist_items": "1",  # trang có nhiều video (vd. bài báo) -> lấy video đầu tiên
        "retries": 5,
        "fragment_retries": 5,
        "socket_timeout": 30,
        "quiet": True,
        "no_warnings": True,
        "noprogress": False,
        "windowsfilenames": True,
        "overwrites": False,
    }
    if ffmpeg_dir:
        opts["ffmpeg_location"] = ffmpeg_dir
    if proxy_hien_tai():
        opts["proxy"] = proxy_hien_tai()   # yt-dlp dùng cho cả giả lập Chrome và ffmpeg (video HLS)
    if PROGRESS_HOOK:
        opts["progress_hooks"] = [PROGRESS_HOOK]
    if COOKIE_GOP:
        opts["cookiefile"] = COOKIE_GOP
    elif COOKIES_FROM_BROWSER:
        opts["cookiesfrombrowser"] = (COOKIES_FROM_BROWSER,)
    return opts


UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36")

# Các nền tảng lớn yt-dlp hỗ trợ sẵn: lỗi ở đây thường do đăng nhập/video bị xóa,
# dò video nhúng không giúp được gì.
NEN_TANG_LON = re.compile(
    r"(youtube\.com|youtu\.be|tiktok\.com|facebook\.com|fb\.watch|instagram\.com|"
    r"twitter\.com|x\.com|pinterest\.|pin\.it)", re.I)

# Iframe nhúng từ các trình phát video phổ biến
HOST_NHUNG = re.compile(
    r"(youtube(-nocookie)?\.com/(embed|watch|shorts)|youtu\.be/|facebook\.com/(plugins/video|.+/videos/)|"
    r"player\.vimeo\.com|vimeo\.com/\d|dailymotion\.com/(embed|video)|dai\.ly/|tiktok\.com/(embed|@)|"
    r"instagram\.com/(p|reel|reels)/|(twitter|x)\.com/.+/status/|cdn\.jwplayer\.com|content\.jwplatform\.com|"
    r"streamable\.com|rumble\.com/embed|ok\.ru/videoembed|vk\.com/video_ext|"
    r"players\.brightcove\.net|fast\.wistia|player\.twitch\.tv|bilibili\.com|tv\.zing|"
    r"embed\.vbear|vbear\.media/embed)", re.I)

QUANG_CAO = re.compile(
    r"(doubleclick|googlesyndication|googleads|imasdk|/ads?/|[_./-]ads?[_./-]|advert|preroll|"
    r"midroll|vast|adserver|adnxs|criteo|teads|outstream|spotx|springserve|"
    r"intro[_-]?logo|bumper|/promo/)", re.I)

DUOI_VIDEO = r"(?:mp4|m3u8|mpd|webm|mov|m4v|flv)"


# File video thuộc GIAO DIỆN trang (không phải nội dung bài đăng), vd. Threads/Instagram/Facebook
RAC = re.compile(r"(/rsrc\.php/|douyinstatic\.com|douyin-pc-web|static\.cdninstagram\.com|static\.xx\.fbcdn\.net|/static/(media|assets)/|"
                 r"byteeffecttos\.com|effectcdn|bytednsdoc\.com|douyin_pc_client|ies\.fe\.effect|"
                 r"placeholder|loading[_-]?anim|/favicon)", re.I)

# Đoạn CHỈ CÓ TIẾNG của video tách hình/tiếng (Douyin, DASH): không bao giờ là video hoàn chỉnh
CHI_CO_TIENG = re.compile(r"/media-audio-|[_/-]audio[_-]?only|mime_type=audio", re.I)


def la_quang_cao(u: str) -> bool:
    return bool(QUANG_CAO.search(u) or RAC.search(u) or CHI_CO_TIENG.search(u))


def kiem_tra_luong(f: Path, ffmpeg_dir):
    """Đọc file bằng ffmpeg. Trả về (có hình, có tiếng, thời lượng giây) hoặc None nếu không kiểm tra được."""
    import subprocess

    if not ffmpeg_dir:
        return None
    exe = str(Path(ffmpeg_dir) / ("ffmpeg.exe" if os.name == "nt" else "ffmpeg"))
    try:
        r = subprocess.run([exe, "-hide_banner", "-i", str(f)], capture_output=True, timeout=60,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        txt = (r.stderr or b"").decode("utf-8", "ignore")
    except Exception:
        return None
    if "Stream #" not in txt:
        return None
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", txt)
    giay = int(m.group(1)) * 3600 + int(m.group(2)) * 60 + float(m.group(3)) if m else 0.0
    # ảnh bìa nhúng trong file nhạc (attached pic) không tính là hình
    co_hinh = any("Video:" in d and "attached pic" not in d for d in txt.splitlines() if "Stream #" in d)
    return co_hinh, bool(re.search(r"Stream #[^\n]*Audio:", txt)), giay


def doc_ma_video(f: Path, ffmpeg_dir):
    """Trả về (mã hình, kiểu điểm ảnh, mã tiếng) của file, vd. ("hevc", "yuv420p", "aac"); None nếu không đọc được."""
    import subprocess

    if not ffmpeg_dir:
        return None
    exe = str(Path(ffmpeg_dir) / ("ffmpeg.exe" if os.name == "nt" else "ffmpeg"))
    try:
        r = subprocess.run([exe, "-hide_banner", "-i", str(f)], capture_output=True, timeout=60,
                           creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        txt = (r.stderr or b"").decode("utf-8", "ignore")
    except Exception:
        return None
    hinh = next((d for d in txt.splitlines() if "Stream #" in d and "Video:" in d and "attached pic" not in d), None)
    if not hinh:
        return None
    m = re.search(r"Video:\s*(\w+)", hinh)
    # kiểu điểm ảnh là mục đầu tiên sau dấu phẩy nằm NGOÀI ngoặc, vd. "hevc (Main) (hvc1 / 0x31637668), yuv420p(tv), ..."
    sau = re.sub(r"\([^()]*\)", "", hinh.split("Video:", 1)[1])
    p = re.search(r",\s*(\w+)", sau)
    a = re.search(r"Stream #[^\n]*Audio:\s*(\w+)", txt)
    return (m.group(1).lower() if m else "", p.group(1).lower() if p else "", a.group(1).lower() if a else "")


def can_chuyen_h264(f: Path, ffmpeg_dir):
    """Premiere và Windows chỉ chắc chắn đọc được MP4 chứa hình H.264 8-bit. Trả về "" nếu file đã ổn,
    "ghep" nếu chỉ cần đổi vỏ file sang MP4, "ma" nếu phải mã hóa lại hình."""
    if not (CHUYEN_H264 and ffmpeg_dir):
        return ""
    ma = doc_ma_video(f, ffmpeg_dir)
    if not ma:
        return ""
    if ma[0] != "h264" or ma[1] not in ("yuv420p", "yuvj420p"):
        return "ma"
    if f.suffix.lower() not in (".mp4", ".mov", ".m4v") or ma[2] not in ("", "aac", "mp3"):
        return "ghep"
    return ""


def dam_bao_h264(f: Path, ffmpeg_dir):
    """Chuyển file về MP4 (H.264 + AAC) nếu cần. Trả về (file, ghi chú). Chuyển hỏng thì GIỮ file gốc."""
    import subprocess

    kieu = can_chuyen_h264(f, ffmpeg_dir)
    if not kieu:
        return f, ""
    ma = doc_ma_video(f, ffmpeg_dir)
    exe = str(Path(ffmpeg_dir) / ("ffmpeg.exe" if os.name == "nt" else "ffmpeg"))
    tam = f.with_name(f.stem + ".dang_chuyen.mp4")
    dich = f.with_suffix(".mp4")
    lenh = [exe, "-hide_banner", "-loglevel", "error", "-y", "-i", str(f), "-map", "0:v:0", "-map", "0:a:0?",
            "-map_metadata", "0"]
    if kieu == "ma":
        print(f"    File tải về dùng mã hình {ma[0].upper()} (Premiere/Windows thường không đọc được) -> đang chuyển sang H.264...")
        lenh += ["-c:v", "libx264", "-preset", "veryfast", "-crf", "18", "-pix_fmt", "yuv420p",
                 "-vf", "scale=trunc(iw/2)*2:trunc(ih/2)*2"]
    else:
        print("    Đổi vỏ file sang MP4 (không mã hóa lại hình)...")
        lenh += ["-c:v", "copy"]
    lenh += (["-c:a", "copy"] if ma[2] in ("aac", "mp3") else ["-c:a", "aac", "-b:a", "192k"])
    lenh += ["-movflags", "+faststart", "-f", "mp4", str(tam)]
    try:
        r = subprocess.run(lenh, capture_output=True, creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        kt = kiem_tra_luong(tam, ffmpeg_dir) if r.returncode == 0 and tam.exists() else None
        if not kt or not kt[0]:
            raise RuntimeError((r.stderr or b"").decode("utf-8", "ignore")[-200:] or "file ra không có hình")
        goc = kiem_tra_luong(f, ffmpeg_dir)
        if goc and goc[2] > 2 and abs(kt[2] - goc[2]) > max(2.0, goc[2] * 0.05):
            raise RuntimeError(f"thời lượng sau khi chuyển lệch ({kt[2]:.0f}s so với {goc[2]:.0f}s)")
        f.unlink()
        tam.replace(dich)
        return dich, ("đã chuyển " + ma[0].upper() + " sang H.264" if kieu == "ma" else "đã đổi vỏ sang MP4")
    except Exception as e:
        try:
            tam.unlink()
        except Exception:
            pass
        print(f"    Không chuyển được sang H.264, giữ file gốc ({str(e).strip()[-160:]})")
        return f, ""


GHI_CHU_NGUON = {}     # link nguồn -> cảnh báo ghi vào báo cáo
NGUON_DA_DUNG = {}     # nguồn dò được -> trang đã dùng nó (phát hiện file giao diện dùng chung)
KICH_THUOC_TOI_THIEU = 60_000   # file dò được nhỏ hơn mức này coi như không phải video


def lam_sach_nguon(u: str):
    """Chuẩn hóa link video của Meta (fbcdn/cdninstagram). Trả về None nếu là đoạn ÂM THANH."""
    if not re.search(r"fbcdn\.net|cdninstagram\.com", u):
        return u
    from urllib.parse import urlsplit, urlunsplit, parse_qsl, urlencode
    import base64

    sp = urlsplit(u)
    q = parse_qsl(sp.query, keep_blank_values=True)
    efg = next((v for k, v in q if k == "efg"), "")
    the = ""
    try:
        the = base64.urlsafe_b64decode(efg + "=" * (-len(efg) % 4)).decode("utf-8", "ignore").lower()
    except Exception:
        pass
    if "audio" in the.split('"vencode_tag"')[-1].split(",")[0]:
        return None
    u2 = urlunsplit(sp._replace(query=urlencode([(k, v) for k, v in q if k not in ("bytestart", "byteend")])))
    if the and "progressive" not in the and "dash" in the:
        GHI_CHU_NGUON[u2] = "CẦN KIỂM TRA: nguồn chỉ có hình, có thể không có tiếng"
    return u2


def chuan_hoa_link(u: str, base: str):
    from html import unescape
    from urllib.parse import urljoin

    u = unescape(u.strip().replace("\\/", "/").replace("\\u0026", "&").replace("\\u002F", "/"))
    if not u or u.startswith(("blob:", "data:", "javascript:")):
        return None
    if u.startswith("//"):
        return "https:" + u
    if re.match(r"^https?://", u, re.I):
        return u
    # dạng "kenh14cdn.com/2025/..../video.mp4" (không có http)
    if re.match(r"^(?:[\w-]+\.)+[a-z]{2,}/", u, re.I) and re.search(r"\." + DUOI_VIDEO + r"\b", u, re.I):
        return "https://" + u
    return urljoin(base, u)


def trich_link_video(html: str, base: str):
    """Tìm link video trong mã HTML của trang. Trả về danh sách theo độ ưu tiên."""
    h = (html.replace("\\/", "/").replace("\\u0026", "&").replace("\\u002F", "/")
         .replace("\\u003D", "=").replace("\\u0025", "%"))
    nhom_meta, nhom_the, nhom_iframe, nhom_regex = [], [], [], []

    # 1. Thẻ meta og:video / twitter:player:stream
    for tag in re.findall(r"<meta\b[^>]+>", h, re.I):
        if re.search(r"""(property|name)\s*=\s*["'](og:video(:url|:secure_url)?|twitter:player:stream)["']""", tag, re.I):
            m = re.search(r"""content\s*=\s*["']([^"']+)""", tag, re.I)
            if m:
                nhom_meta.append(m.group(1))
    # JSON-LD VideoObject
    nhom_meta += re.findall(r'"contentUrl"\s*:\s*"([^"]+)"', h)
    nhom_meta += [u for u in re.findall(r'"embedUrl"\s*:\s*"([^"]+)"', h)]

    # 2. <video src>, <source src>, và các thuộc tính data-* của trình phát
    nhom_the += re.findall(r"""<(?:video|source)\b[^>]*?\ssrc\s*=\s*["']([^"']+)""", h, re.I)
    for val in re.findall(
            r"""\sdata-(?:vid|src|video|video-src|videosrc|file|url|mp4|hls|stream|source|link)\s*=\s*["']([^"']+)""",
            h, re.I):
        if re.search(r"\." + DUOI_VIDEO + r"(\b|$)", val, re.I) or HOST_NHUNG.search(val):
            nhom_the.append(val)
    # biến JS kiểu file: "...mp4", src: '...m3u8'
    nhom_the += re.findall(
        r"""(?:file|src|source|url|hls|mp4|videoUrl|video_url|playUrl)["']?\s*[:=]\s*["']([^"']+\.""" + DUOI_VIDEO + r"""[^"']*)""",
        h, re.I)

    # 3. iframe nhúng
    for src in re.findall(r"""<iframe\b[^>]*?\s(?:data-)?src\s*=\s*["']([^"']+)""", h, re.I):
        if HOST_NHUNG.search(src):
            nhom_iframe.append(src)

    # 3b. Bài X/Instagram/TikTok nhúng dạng <blockquote> (báo hay nhúng tweet có video kiểu này)
    for khoi in re.findall(r"<blockquote\b[^>]*class\s*=\s*[\"'][^\"']*(?:twitter-tweet|instagram-media|tiktok-embed)"
                           r"[^\"']*[\"'][^>]*>.*?</blockquote>", h, re.I | re.S):
        ung = (re.findall(r"""(?:cite|data-instgrm-permalink)\s*=\s*["']([^"']+)""", khoi, re.I)
               + re.findall(r"""href\s*=\s*["'](https?://(?:www\.)?(?:twitter|x)\.com/[^"'/]+/status/\d+)""", khoi, re.I))
        for u in ung:
            u = u.split("?")[0]
            if HOST_NHUNG.search(u):
                nhom_iframe.append(u)

    # 3c. Trang phát video bằng trình phát Dailymotion và ghi mã video trong dữ liệu trang chứ không có iframe sẵn
    #     (vd. báo của nhóm RCS: video.corriere.it). "privateId" đứng trước vì video của báo thường để riêng tư,
    #     chỉ mở được bằng mã riêng mà chính trang đó dùng để phát.
    if "dailymotion" in h.lower():
        for khoa in ("privateId", "dmId"):
            for ma in re.findall(r'"%s"\s*:\s*"([A-Za-z0-9]{6,24})"' % khoa, h):
                nhom_iframe.append("https://www.dailymotion.com/video/" + ma)
        for ma in re.findall(r"""geo\.dailymotion\.com/player/\w+\.html\?(?:[^"'\s<>]*&)?video=([A-Za-z0-9]{6,24})""", h):
            nhom_iframe.append("https://www.dailymotion.com/video/" + ma)

    # 4. Quét toàn bộ trang tìm đường dẫn video
    nhom_regex += re.findall(
        r"""((?:https?:)?//[^\s"'<>()\\]+?\.""" + DUOI_VIDEO + r"""(?:\?[^\s"'<>()\\]*)?)(?=["'\s<>)\\]|$)""",
        h, re.I)
    nhom_regex += re.findall(
        r"""(?<![\w/.:-])((?:[\w-]+\.)+[a-z]{2,}/[^\s"'<>()\\]+?\.(?:mp4|m3u8))(?=["'\s<>)\\?]|$)""",
        h, re.I)

    out = []
    for u in nhom_meta + nhom_the + nhom_iframe + nhom_regex:
        u2 = chuan_hoa_link(u, base)
        if not u2 or la_quang_cao(u2) or u2 in out:
            continue
        if u2.rstrip("/") == base.rstrip("/"):
            continue
        if not (re.search(r"\." + DUOI_VIDEO + r"(\?|#|$)", u2, re.I) or HOST_NHUNG.search(u2)):
            continue
        out.append(u2)
    return out


UA_DI_DONG = ("Mozilla/5.0 (iPhone; CPU iPhone OS 17_5 like Mac OS X) AppleWebKit/605.1.15 "
              "(KHTML, like Gecko) Version/17.5 Mobile/15E148 Safari/604.1")


def do_douyin(video_id: str):
    """Douyin không cần cookie: trang chia sẻ bản di động có sẵn link phát video."""
    html = tai_html(f"https://www.iesdouyin.com/share/video/{video_id}/", ua=UA_DI_DONG)
    h = html.replace("\\u002F", "/").replace("\\/", "/").replace("\\u0026", "&")
    out = []
    ung_vien = (re.findall(r'"play_addr"\s*:\s*\{.*?"url_list"\s*:\s*\[\s*"([^"]+)"', h, re.S)
                + re.findall(r'(https?://[^"\s\\]+/aweme/v1/playwm?/\?[^"\s\\]+)', h))
    for uri in re.findall(r'"play_addr"\s*:\s*\{\s*"uri"\s*:\s*"([\w-]+)"', h):
        ung_vien.append(f"https://aweme.snssdk.com/aweme/v1/play/?video_id={uri}&ratio=1080p&line=0")
        ung_vien.append(f"https://www.iesdouyin.com/aweme/v1/play/?video_id={uri}&ratio=1080p&line=0")
    for u in ung_vien:
        u = u.replace("/playwm/", "/play/")   # bản không đóng dấu
        if u.startswith("//"):
            u = "https:" + u
        if u.startswith("http") and u not in out:
            out.append(u)
    if not out:
        raise RuntimeError(f"trang chia sẻ không có link phát (dài {len(html)} ký tự"
                           + (", bị chuyển sang trang xác minh" if "captcha" in h.lower() or "verify" in h.lower() else "") + ")")
    return out


def douyin_id(url: str):
    """Mã video thật của link Douyin (link rút gọn v.douyin.com được mở theo để lấy mã)."""
    import urllib.request

    mau = r"(?:/(?:video|note|slides)/|[?&](?:modal_id|aweme_id|vid)=)(\d{15,})"
    m = re.search(mau, url)
    if m:
        return m.group(1)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA_DI_DONG})
        with _mo_url(req, 25) as r:
            m = re.search(mau, r.geturl())
            return m.group(1) if m else None
    except Exception:
        return None


def _douyin_link_tu_du_lieu(video: dict):
    """Lấy danh sách link phát từ dữ liệu 'video' của Douyin, ưu tiên: file mp4 đủ hình+tiếng,
    mã H.264 (mở được ở mọi phần mềm dựng), chất lượng cao nhất."""
    def ds_link(o):
        if isinstance(o, str):
            return [o]
        if isinstance(o, dict):
            return ds_link(o.get("url_list") or o.get("urlList") or o.get("src") or [])
        if isinstance(o, list):
            return [x for i in o for x in ds_link(i)]
        return []

    muc = []
    for br in (video.get("bit_rate") or video.get("bitRateList") or []):
        if not isinstance(br, dict) or str(br.get("format") or "mp4").lower() == "dash":
            continue
        h265 = 1 if (br.get("is_h265") or br.get("isH265")) else 0
        toc_do = br.get("bit_rate") or br.get("bitRate") or 0
        link = ds_link(br.get("play_addr") or br.get("playAddr") or br.get("playApi"))
        if link:
            muc.append((h265, -int(toc_do or 0), link))
    muc.sort(key=lambda x: x[:2])
    out = [u for _, _, ds in muc for u in ds]
    for khoa in ("play_addr_h264", "play_addr", "playAddr", "playApi", "play_addr_265"):
        out += ds_link(video.get(khoa))
    kq = []
    for u in out:
        if u.startswith("//"):
            u = "https:" + u
        if u.startswith("http") and not CHI_CO_TIENG.search(u) and "/media-video-" not in u and u not in kq:
            kq.append(u)
    return kq


_DOUYIN_PHIEN = {}   # phiên trình duyệt của lần lấy dữ liệu Douyin thành công gần nhất


def do_douyin_trinh_duyet(video_id: str):
    """Mở douyin.com/video/<mã> bằng trình duyệt ẩn và đọc dữ liệu CHÍNH THỨC của đúng video đó
    (không bắt bừa video đang phát trên trang, vì trang Douyin còn phát cả video gợi ý và clip giao diện).
    Trả về (danh sách link phát, thời lượng giây)."""
    if TRANG_THAI_TD != "ok":
        raise RuntimeError(f"trình duyệt ẩn không chạy ({TRANG_THAI_TD})")
    import json
    from urllib.parse import unquote
    from playwright.sync_api import sync_playwright

    nhan = []      # dữ liệu aweme_detail đúng mã
    tu_choi = []   # lý do Douyin trả về khi không cho xem

    def on_response(resp):
        try:
            if "/aweme/detail/" not in resp.url:
                return
            d = resp.json()
            ad = d.get("aweme_detail") if isinstance(d, dict) else None
            if ad and str(ad.get("aweme_id")) == video_id:
                nhan.append(ad)
            elif isinstance(d, dict):
                fd = d.get("filter_detail") or {}
                ly_do = fd.get("detail_msg") or fd.get("notice") or d.get("status_msg") or ""
                if ly_do:
                    tu_choi.append(str(ly_do)[:80])
        except Exception:
            pass   # lần gọi đầu thường trả về rỗng, trang sẽ tự gọi lại

    def tim_trong_trang(o, sau=0):
        if sau > 14:
            return None
        if isinstance(o, dict):
            if str(o.get("awemeId") or o.get("aweme_id") or "") == video_id and isinstance(o.get("video"), dict):
                return o
            o = list(o.values())
        if isinstance(o, list):
            for v in o:
                r = tim_trong_trang(v, sau + 1)
                if r:
                    return r
        return None

    with sync_playwright() as p:
        b = mo_trinh_duyet(p)
        if b is None:
            raise RuntimeError("không mở được trình duyệt ẩn")
        try:
            ctx = page = None
            for lan in range(4):
                # Douyin thỉnh thoảng trả dữ liệu rỗng: thử tải lại trang, rồi mở phiên trình duyệt mới
                if lan % 2 == 0:
                    if ctx is not None:
                        ctx.close()
                        import time
                        time.sleep(4)
                    # Dùng lại phiên (cookie) của lần lấy được gần nhất: Douyin ít trả rỗng hơn hẳn
                    ctx = b.new_context(user_agent=UA, locale="zh-CN", viewport={"width": 1280, "height": 900},
                                        storage_state=_DOUYIN_PHIEN.get("p") if lan == 0 else None)
                    page = ctx.new_page()
                    page.on("response", on_response)
                try:
                    if lan % 2 == 0:
                        page.goto(f"https://www.douyin.com/video/{video_id}", wait_until="domcontentloaded", timeout=45000)
                    else:
                        page.reload(wait_until="domcontentloaded", timeout=45000)   # đã có cookie mới -> thử lại
                except Exception:
                    pass
                for _ in range(24):
                    if nhan:
                        break
                    page.wait_for_timeout(500)
                if nhan:
                    break
                try:
                    m = re.search(r'<script id="RENDER_DATA"[^>]*>([^<]+)</script>', page.content())
                    a = tim_trong_trang(json.loads(unquote(m.group(1)))) if m else None
                    if a:
                        nhan.append(a)
                        break
                except Exception:
                    pass
            try:
                if nhan:
                    _DOUYIN_PHIEN["p"] = ctx.storage_state()
                else:
                    _DOUYIN_PHIEN.pop("p", None)
            except Exception:
                pass
        finally:
            b.close()
    if not nhan:
        raise RuntimeError("Douyin không trả dữ liệu của video này"
                           + (f" ({tu_choi[-1]})" if tu_choi else " (video có thể đã bị xóa/ẩn, hoặc Douyin đang chặn)"))
    ad = nhan[0]
    video = ad.get("video") or {}
    link = _douyin_link_tu_du_lieu(video)
    if not link:
        if ad.get("images"):
            raise RuntimeError("bài Douyin này là bộ ẢNH, không phải video")
        raise RuntimeError("dữ liệu Douyin không có link phát")
    giay = float(video.get("duration") or 0)
    return link, (giay / 1000 if giay > 1000 else giay)


def tai_douyin(video_id: str, outdir: Path, name: str, ffmpeg_dir):
    """Tải đúng video Douyin theo mã, rồi kiểm tra file: phải có hình và dài đúng như Douyin báo."""
    link, giay = do_douyin_trinh_duyet(video_id)
    dich = outdir / f"{name}.mp4"
    loi = []
    for u in link[:6]:
        don_rac(outdir, name)
        try:
            tai_truc_tiep(u, dich, UA, "https://www.douyin.com/")
        except Exception as e:
            loi.append(str(e)[:70])
            continue
        kt = kiem_tra_luong(dich, ffmpeg_dir)
        if kt:
            co_hinh, co_tieng, dai = kt
            if not co_hinh or (giay and dai and abs(dai - giay) > max(2.0, giay * 0.05)):
                dich.unlink(missing_ok=True)
                loi.append(f"file không khớp (hình: {co_hinh}, dài {dai:.0f}s, Douyin báo {giay:.0f}s)")
                continue
            if not co_tieng:
                return dich, f"Douyin: dữ liệu chính thức, mã {video_id} (CẦN KIỂM TRA: file không có tiếng)"
        return dich, f"Douyin: dữ liệu chính thức, mã {video_id}"
    raise RuntimeError("có link phát nhưng không tải được: " + "; ".join(dict.fromkeys(loi)))


def tai_truc_tiep(u: str, dest: Path, ua: str, referer: str):
    """Tải thẳng 1 file video (không qua yt-dlp)."""
    import urllib.request

    req = urllib.request.Request(u, headers={"User-Agent": ua, "Referer": referer, "Accept": "*/*"})
    tam = dest.with_name(dest.name + ".part")
    with _mo_url(req, 60) as r:
        ct = (r.headers.get("Content-Type") or "").lower()
        if not (ct.startswith("video/") or "octet-stream" in ct):
            raise RuntimeError(f"máy chủ không trả về video ({ct or 'không rõ loại'})")
        tong = int(r.headers.get("Content-Length") or 0)
        da = 0
        with open(tam, "wb") as fh:
            while True:
                khoi = r.read(256 * 1024)
                if not khoi:
                    break
                fh.write(khoi)
                da += len(khoi)
                if PROGRESS_HOOK:
                    PROGRESS_HOOK({"status": "downloading", "downloaded_bytes": da, "total_bytes": tong or None})
    if da < KICH_THUOC_TOI_THIEU:
        tam.unlink(missing_ok=True)
        raise RuntimeError(f"file quá nhỏ ({da} byte)")
    tam.replace(dest)
    return dest


_KHOA_TG = threading.Lock()
_TG_LAN_CUOI = [0.0]


def tai_tiktok_trung_gian(url: str, outdir: Path, name: str, ffmpeg_dir):
    """TikTok bắt đăng nhập ("may not be comfortable for some audiences"): nhờ dịch vụ trung gian tikwm.com lấy hộ,
    giống cách các trang kiểu snaptik làm. CHỈ gửi đi địa chỉ video. Trả về (file, cách lấy)."""
    import json
    import time
    import urllib.parse
    import urllib.request

    with _KHOA_TG:   # cổng miễn phí của họ giới hạn 1 lượt/giây
        cho = 1.3 - (time.time() - _TG_LAN_CUOI[0])
        if cho > 0:
            time.sleep(cho)
        try:
            req = urllib.request.Request(
                "https://www.tikwm.com/api/?" + urllib.parse.urlencode({"url": url, "hd": 1}),
                headers={"User-Agent": UA, "Accept": "application/json"})
            with _mo_url(req, 40) as r:
                d = json.loads(r.read().decode("utf-8", "ignore"))
        finally:
            _TG_LAN_CUOI[0] = time.time()
    dl = d.get("data") if isinstance(d, dict) else None
    if not isinstance(dl, dict) or d.get("code") != 0:
        raise RuntimeError(str((d or {}).get("msg") or "không có dữ liệu")[:120])
    m = re.search(r"/(?:video|photo|v)/(\d{15,})", url)
    if m and str(dl.get("id") or "") != m.group(1):
        raise RuntimeError("dịch vụ trả về video khác mã với link")
    if dl.get("images") and not (dl.get("hdplay") or dl.get("play")):
        raise RuntimeError("đây là bài ảnh, không phải video")
    dich = outdir / f"{name}.mp4"
    loi = []
    for u in dict.fromkeys(x for x in (dl.get("hdplay"), dl.get("play")) if x):
        if u.startswith("/"):
            u = "https://www.tikwm.com" + u
        don_rac(outdir, name)
        try:
            tai_truc_tiep(u, dich, UA, "https://www.tikwm.com/" if "tikwm.com" in u else "https://www.tiktok.com/")
            kt = kiem_tra_luong(dich, ffmpeg_dir)
            if kt and not kt[0]:
                raise RuntimeError("file không có hình")
            dai = float(dl.get("duration") or 0)
            if kt and dai > 2 and abs(kt[2] - dai) > max(3.0, dai * 0.1):
                raise RuntimeError(f"thời lượng lệch ({kt[2]:.0f}s so với {dai:.0f}s)")
            return dich, "dịch vụ trung gian tikwm.com (TikTok bắt đăng nhập)"
        except Exception as e:
            dich.unlink(missing_ok=True)
            loi.append(re.sub(r"\s+", " ", str(e))[:80])
    raise RuntimeError("có link nhưng không tải được: " + "; ".join(dict.fromkeys(loi)) if loi else "không có link video")


def tai_x_trung_gian(url: str, outdir: Path, name: str, ffmpeg_dir):
    """X/Twitter không cho khách xem video của tweet (yt-dlp báo "No video could be found" hoặc đòi đăng nhập):
    hỏi dịch vụ mã nguồn mở FxTwitter (api.fxtwitter.com, cổng công khai có tài liệu). CHỈ gửi mã tweet.
    Tweet không có video nhưng trích dẫn một tweet có video thì lấy video của tweet được trích dẫn."""
    import json
    import urllib.request

    m = re.search(r"(?:twitter|x)\.com/[^/]+/status(?:es)?/(\d+)", url, re.I)
    if not m:
        raise RuntimeError("không đọc được mã tweet từ link")
    req = urllib.request.Request(
        "https://api.fxtwitter.com/2/status/" + m.group(1),
        headers={"User-Agent": f"TaiVideoExcel/{PHIEN_BAN} (bot; video downloader)", "Accept": "application/json"})
    with _mo_url(req, 40) as r:
        d = json.loads(r.read().decode("utf-8", "ignore"))
    tw = d.get("status") if isinstance(d, dict) else None
    if not isinstance(tw, dict):
        raise RuntimeError(str((d or {}).get("message") or "không có dữ liệu")[:120])
    nguon = ""
    video = ((tw.get("media") or {}).get("videos") or [])
    if not video and isinstance(tw.get("quote"), dict):
        video = ((tw["quote"].get("media") or {}).get("videos") or [])
        nguon = " - video của tweet được trích dẫn"
    if not video:
        raise RuntimeError("tweet này không có video")
    v = video[0]

    def cao(u):
        k = re.search(r"/(\d+)x(\d+)/", u or "")
        return min(int(k.group(1)), int(k.group(2))) if k else 0

    ds = [f for f in (v.get("formats") or []) if f.get("container") == "mp4" and f.get("url")]
    if MAX_HEIGHT and any(0 < cao(f["url"]) <= MAX_HEIGHT for f in ds):
        ds = [f for f in ds if 0 < cao(f["url"]) <= MAX_HEIGHT]
    ds.sort(key=lambda f: int(f.get("bitrate") or 0), reverse=True)
    ung_vien = list(dict.fromkeys([f["url"] for f in ds] + ([v["url"]] if str(v.get("url") or "").split("?")[0].endswith(".mp4") else [])))
    dich = outdir / f"{name}.mp4"
    loi = []
    for u in ung_vien[:3]:
        don_rac(outdir, name)
        try:
            tai_truc_tiep(u, dich, UA, "https://x.com/")
            kt = kiem_tra_luong(dich, ffmpeg_dir)
            if kt and not kt[0]:
                raise RuntimeError("file không có hình")
            dai = float(v.get("duration") or 0)
            if kt and dai > 2 and abs(kt[2] - dai) > max(3.0, dai * 0.1):
                raise RuntimeError(f"thời lượng lệch ({kt[2]:.0f}s so với {dai:.0f}s)")
            return dich, "dịch vụ trung gian FxTwitter (X không cho khách xem)" + nguon
        except Exception as e:
            dich.unlink(missing_ok=True)
            loi.append(" ".join(str(e).split())[:80])
    raise RuntimeError("có link nhưng không tải được: " + "; ".join(dict.fromkeys(loi)) if loi else "không có bản mp4")


def tai_html(url: str, ua: str = None) -> str:
    import urllib.request

    try:
        req = urllib.request.Request(url, headers={
            "User-Agent": ua or UA, "Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8",
            "Accept": "text/html,application/xhtml+xml,*/*;q=0.8"})
        with _mo_url(req, 30) as r:
            raw = r.read()
            cs = r.headers.get_content_charset() or "utf-8"
        return raw.decode(cs, errors="ignore")
    except Exception as e1:
        # Trang chặn bot (Cloudflare, 403...): thử lại bằng curl_cffi giả lập Chrome thật
        try:
            from curl_cffi import requests as creq
        except ImportError:
            raise e1
        p = proxy_hien_tai()
        r = creq.get(url, impersonate="chrome", timeout=30, headers={"Accept-Language": "vi-VN,vi;q=0.9,en;q=0.8"},
                     proxies={"http": p, "https": p} if p else None)
        if r.status_code >= 400:
            raise RuntimeError(f"HTTP {r.status_code}")
        return r.text


def do_html(url: str):
    return trich_link_video(tai_html(url), url)


def luat_dns(url: str):
    """Một số mạng không phân giải được tên 'www.<trang>' dù tên gốc '<trang>' vẫn được (vd. www.reddit.com).
    Khi đó trả về tham số để trình duyệt ẩn dùng địa chỉ của tên gốc cho 'www.<trang>'."""
    import socket
    from urllib.parse import urlsplit

    may = (urlsplit(url).hostname or "").lower()
    if not may.startswith("www.") or proxy_hien_tai():   # qua proxy thì proxy tự phân giải tên
        return []
    try:
        socket.gethostbyname(may)
        return []
    except OSError:
        pass
    try:
        return [f"--host-resolver-rules=MAP {may} {socket.gethostbyname(may[4:])}"]
    except OSError:
        return []


def mo_trinh_duyet(p, them=None):
    args = ["--autoplay-policy=no-user-gesture-required", "--mute-audio",
            "--disable-blink-features=AutomationControlled"] + list(them or [])
    px = _proxy_cho_trinh_duyet()
    for kw in ({"channel": "msedge"}, {"channel": "chrome"}, {}):
        if px:
            kw = dict(kw, proxy=px)
        try:
            return p.chromium.launch(headless=True, args=args, **kw)
        except Exception:
            continue
    return None


def dam_bao_goi():
    """Tự cài các thư viện còn thiếu (kể cả khi chạy thẳng tai_video.py, không qua CHAY_TOOL.bat)."""
    import importlib
    import subprocess

    print(f"HT CREATIVE AGENT phiên bản {PHIEN_BAN} | Python {sys.version.split()[0]}")
    for mod, goi in (("yt_dlp", "yt-dlp[default]"), ("openpyxl", "openpyxl"), ("imageio_ffmpeg", "imageio-ffmpeg"),
                     ("curl_cffi", "yt-dlp[default,curl-cffi]"), ("playwright", "playwright")):
        try:
            importlib.import_module(mod)
            continue
        except ImportError:
            pass
        print(f"Thiếu thư viện {mod} -> đang cài ({goi})...")
        try:
            r = subprocess.run([sys.executable, "-m", "pip", "install", "--disable-pip-version-check", "-q", goi],
                               capture_output=True, text=True, timeout=900,
                               creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
            importlib.invalidate_caches()
            if r.returncode == 0:
                print("    Đã cài xong.")
            else:
                dong = [x for x in (r.stderr or r.stdout or "").strip().splitlines() if x.strip()][-3:]
                print("    CÀI THẤT BẠI: " + " | ".join(dong)[:400])
                if mod in ("playwright", "curl_cffi"):
                    print("    (Nếu Python quá mới chưa có bản cài sẵn: cài Python 3.12 rồi chạy lại.)")
        except Exception as e:
            print(f"    CÀI THẤT BẠI: {e}")


def kiem_tra_trinh_duyet() -> str:
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return "chưa cài playwright (chạy: python -m pip install playwright)"
    try:
        with sync_playwright() as p:
            b = mo_trinh_duyet(p)
            if b is None:
                return "không mở được Edge/Chrome (chạy: python -m playwright install chromium)"
            b.close()
    except Exception as e:
        return ("lỗi playwright: " + str(e).replace("\n", " "))[:160]
    return "ok"


def do_trinh_duyet(url: str):
    """Mở trang bằng Edge/Chrome ẩn, bấm phát và bắt link video trong lưu lượng mạng.
    Dùng cho trang tải video bằng JavaScript. Cần: pip install playwright"""
    if TRANG_THAI_TD != "ok":
        return []
    from playwright.sync_api import sync_playwright

    found = []

    def them(u):
        u = lam_sach_nguon(u) if u else u
        if u and u.startswith("http") and not la_quang_cao(u) and u not in found:
            if not re.search(r"\.(ts|m4s|aac|m4a)(\?|$)", u.split("#")[0], re.I):
                found.append(u)

    def on_response(resp):
        try:
            u = resp.url
            ct = (resp.headers.get("content-type") or "").lower()
            if (re.search(r"\.(m3u8|mpd|mp4|webm|mov|m4v)(\?|$)", u, re.I)
                    or ct.startswith("video/") or "mpegurl" in ct or "dash+xml" in ct):
                them(u)
        except Exception:
            pass

    js_lay_nguon = """() => {
      const out = [];
      document.querySelectorAll('video').forEach(v => {
        try { v.muted = true; const p = v.play(); if (p) p.catch(()=>{}); } catch (e) {}
        if (v.currentSrc) out.push(v.currentSrc);
        if (v.src) out.push(v.src);
        v.querySelectorAll('source').forEach(s => s.src && out.push(s.src));
      });
      document.querySelectorAll('iframe').forEach(f => f.src && out.push(f.src));
      return out;
    }"""

    with sync_playwright() as p:
        browser = mo_trinh_duyet(p, luat_dns(url))
        if browser is None:
            return []
        try:
            ctx = browser.new_context(user_agent=UA, locale="vi-VN", viewport={"width": 1280, "height": 900})
            if proxy_hien_tai():
                # Trang báo tải hàng trăm ảnh/quảng cáo; qua proxy tính tiền theo dung lượng thì bỏ ảnh và phông chữ
                ctx.route("**/*", lambda r: r.abort() if r.request.resource_type in ("image", "font")
                          else r.continue_())
            page = ctx.new_page()
            page.on("response", on_response)
            try:
                page.goto(url, wait_until="domcontentloaded", timeout=45000)
            except Exception:
                pass
            page.wait_for_timeout(3000)
            for _ in range(5):
                page.mouse.wheel(0, 1000)
                page.wait_for_timeout(600)
            page.evaluate("window.scrollTo(0, 0)")
            # Bấm nút phát của các trình phát phổ biến
            for sel in ["video", ".vjs-big-play-button", ".jw-icon-display", ".jw-display-icon-container",
                        ".plyr__control--overlaid", ".ytp-large-play-button", "[class*='play-button']",
                        "[class*='btn-play']", "[class*='videoplay']", "[class*='icon-play']"]:
                try:
                    el = page.query_selector(sel)
                    if el and el.is_visible():
                        el.scroll_into_view_if_needed(timeout=2000)
                        el.click(timeout=2000)
                        break
                except Exception:
                    continue
            try:
                for u in page.evaluate(js_lay_nguon) or []:
                    if u.startswith("http") and (re.search(r"\." + DUOI_VIDEO + r"(\?|$)", u, re.I)
                                                 or HOST_NHUNG.search(u)):
                        them(u)
            except Exception:
                pass
            page.wait_for_timeout(7000)
            for fr in page.frames:
                if HOST_NHUNG.search(fr.url or ""):
                    them(fr.url)
            try:
                luu_cookie_tu_dong(ctx.cookies())
            except Exception:
                pass
            # Dò thêm trong HTML đã chạy JavaScript
            try:
                for u in trich_link_video(page.content(), url):
                    them(u)
            except Exception:
                pass
        finally:
            browser.close()
    # Ưu tiên playlist HLS/DASH và file mp4 hơn iframe; nguồn chỉ có hình xếp sau
    found.sort(key=lambda u: 3 if u in GHI_CHU_NGUON else 0 if re.search(r"\.(m3u8|mpd)(\?|$)", u, re.I)
               else 1 if re.search(r"\." + DUOI_VIDEO + r"(\?|$)", u, re.I) else 2)
    return found


def do_reddit(url: str):
    """Reddit: mở <bài đăng>.json bằng trình duyệt ẩn để lấy đúng video của bài.
    Trình duyệt dùng DNS riêng nên vẫn vào được khi DNS của Windows bị chặn."""
    if TRANG_THAI_TD != "ok":
        return []
    import json
    from playwright.sync_api import sync_playwright

    m = re.search(r"https?://(?:[\w-]+\.)?reddit\.com(/r/[^/]+/comments/[^/?#]+)", url, re.I)
    if not m:
        return []
    out = []
    with sync_playwright() as p:
        # Nhiều mạng không phân giải được tên www.reddit.com (reddit.com thì được) -> xem luat_dns
        b = mo_trinh_duyet(p, luat_dns("https://www.reddit.com/"))
        if b is None:
            return []
        try:
            page = b.new_context(user_agent=UA, locale="en-US").new_page()
            data = None
            loi = ""
            for may_chu in ("www.reddit.com", "old.reddit.com"):
                try:
                    page.goto(f"https://{may_chu}{m.group(1)}/.json?raw_json=1", wait_until="domcontentloaded",
                              timeout=45000)
                    data = json.loads(page.evaluate("document.body.innerText"))
                    break
                except Exception as e:
                    loi = str(e)
            if data is None:
                raise RuntimeError(loi)
            bai = data[0]["data"]["children"][0]["data"]
            for nguon in [bai] + list(bai.get("crosspost_parent_list") or []):
                for khoa in ("secure_media", "media"):
                    v = ((nguon.get(khoa) or {}).get("reddit_video") or {})
                    for k in ("hls_url", "fallback_url"):
                        if v.get(k) and v[k] not in out:
                            out.append(v[k])
                ngoai = nguon.get("url_overridden_by_dest") or ""
                if ngoai.startswith("http") and "reddit.com" not in ngoai and not re.search(r"\.(jpe?g|png|gif|webp)(\?|$)", ngoai, re.I):
                    if ngoai not in out:
                        out.append(ngoai)
        except Exception as e:
            print(f"    (Reddit .json không đọc được, chuyển sang đọc trang bài đăng: {str(e)[:60]})")
        try:
            if not out:
                # Reddit hay chặn .json. Đọc trình phát nằm TRONG đúng khung bài đăng (theo mã bài),
                # không lấy video của các bài gợi ý bên cạnh.
                ma_bai = m.group(1).rstrip("/").split("/")[-1]
                page.goto(f"https://www.reddit.com{m.group(1)}/", wait_until="domcontentloaded", timeout=45000)
                page.wait_for_timeout(2500)
                kq = page.evaluate("""(id) => {
                  const p = document.querySelector('shreddit-post[id="t3_' + id + '"]');
                  if (!p) return null;
                  const out = [];
                  p.querySelectorAll('shreddit-player, shreddit-player-2, video, source').forEach(
                    e => { const s = e.getAttribute('src'); if (s) out.push(s); });
                  const h = p.getAttribute('content-href') || '';
                  if (/^https?:/.test(h)) out.push(h);
                  return out;
                }""", ma_bai)
                if kq is None:
                    print("    (Reddit: không mở được bài đăng — có thể đã bị gỡ hoặc cần đăng nhập)")
                for u in kq or []:
                    if u.startswith("http") and u not in out and not re.search(
                            r"^https?://v\.redd\.it/\w+/?$|\.(jpe?g|png|gif|webp)(\?|$)|reddit\.com/", u, re.I):
                        out.append(u)
            if not out:
                print("    (Reddit: bài đăng này không có video — có thể là ảnh hoặc đã bị gỡ)")
        except Exception as e:
            print(f"    (Reddit: đọc trang bài đăng lỗi: {str(e)[:100]})")
        finally:
            b.close()
    return out


def ytdlp_tai(url, outdir: Path, name, ffmpeg_dir, referer=None, ep_mp4=False, cookiefile=None,
              impersonate=False, ua=None):
    import yt_dlp

    opts = tao_options(outdir, name, ffmpeg_dir)
    if cookiefile:
        opts["cookiefile"] = cookiefile
        opts.pop("cookiesfrombrowser", None)
    if impersonate:
        # Giả lập Chrome thật (cần gói curl_cffi) để qua Cloudflare / lỗi SSL / 403
        from yt_dlp.networking.impersonate import ImpersonateTarget
        opts["impersonate"] = ImpersonateTarget("chrome")
        opts["extractor_args"] = {"generic": {"impersonate": ["chrome"]}}
    if referer:
        opts["http_headers"] = {"Referer": referer, "User-Agent": ua or UA}
    if ep_mp4:
        opts["outtmpl"] = str(outdir / f"{name}.mp4")
    if opts.get("cookiefile"):
        opts["cookiefile"] = _ban_sao_cookie(opts["cookiefile"])
    da_bat = False
    if ep_mp4:
        # Cho phép đuôi file lạ (vd. Snapchat), file vẫn được lưu là .mp4.
        # Đây là thay đổi chung cho cả chương trình nên phải đếm số luồng đang cần, luồng cuối mới trả lại hàm gốc.
        try:
            from yt_dlp.utils import _UnsafeExtensionError as U
            with _KHOA_DUOI:
                if _DEM_DUOI[0] == 0:
                    _DEM_DUOI[1] = U.__dict__.get("sanitize_extension")
                    U.sanitize_extension = classmethod(lambda cls, ext, prepend=False: ext)
                _DEM_DUOI[0] += 1
                da_bat = True
        except Exception:
            pass
    try:
        with yt_dlp.YoutubeDL(opts) as ydl:
            ydl.download([url])
    finally:
        if da_bat:
            from yt_dlp.utils import _UnsafeExtensionError as U
            with _KHOA_DUOI:
                _DEM_DUOI[0] -= 1
                if _DEM_DUOI[0] == 0 and _DEM_DUOI[1] is not None:
                    U.sanitize_extension = _DEM_DUOI[1]
    f = da_tai(outdir, name)
    if not f:
        raise RuntimeError("Không tìm thấy file sau khi tải")
    return f


def tai_mot(item, outdir: Path, ffmpeg_dir):
    """Tải một link rồi chuyển file về H.264 nếu cần (để dùng được trong Premiere)."""
    f, cach = _tai_mot_proxy(item, outdir, ffmpeg_dir)
    f, ghi_chu = dam_bao_h264(f, ffmpeg_dir)
    return f, (f"{cach} [{ghi_chu}]" if ghi_chu else cach)


def _tai_mot_proxy(item, outdir: Path, ffmpeg_dir):
    """Tải một link, có tính đến proxy: tắt -> đi thẳng; "luon" -> luôn qua proxy;
    "khi_loi" -> đi thẳng trước, chỉ khi lỗi do mạng/IP mới thử lại qua proxy (tiết kiệm dung lượng proxy)."""
    ds = PROXY_DS if PROXY_CHE_DO != "tat" else []
    _TL.proxy = None
    if not ds:
        return _tai_mot(item, outdir, ffmpeg_dir)
    p = _chon_proxy()
    try:
        if PROXY_CHE_DO == "luon":
            _TL.proxy = p
            f, cach = _tai_mot(item, outdir, ffmpeg_dir)
            return f, f"{cach} [qua proxy {che_proxy(p)}]"
        try:
            return _tai_mot(item, outdir, ffmpeg_dir)
        except Exception as e:
            if not _LOI_MANG.search(str(e)):
                raise
            print(f"    Tải thẳng không được (lỗi mạng/IP) -> thử lại qua proxy {che_proxy(p)}...")
            don_rac(outdir, item["name"])
            _TL.proxy = p
            try:
                f, cach = _tai_mot(item, outdir, ffmpeg_dir)
                return f, f"{cach} [qua proxy {che_proxy(p)}]"
            except Exception as e2:
                raise RuntimeError(f"{e2} [đã thử cả đi thẳng lẫn qua proxy {che_proxy(p)}]")
    finally:
        _TL.proxy = None


def _tai_mot(item, outdir: Path, ffmpeg_dir):
    """Trả về (file, cách lấy). Thử lần lượt:
    1) yt-dlp trực tiếp
    2) các cách sửa riêng: link Douyin rút gọn, đuôi file lạ, giả lập Chrome khi bị chặn
    3) dò video nhúng trong HTML
    4) dò bằng trình duyệt ẩn (+ tự lấy cookie cho trang đòi 'fresh cookies' như Douyin)."""
    url, name = item["link"], item["name"]
    url = re.sub(r"^(https?://)linkedin\.com/", r"\1www.linkedin.com/", url, flags=re.I)
    if re.search(r"reddit\.com/r/[^/]+/?(\?.*)?$", url, re.I):
        raise RuntimeError("Đây là link NHÓM Reddit, không phải link bài đăng có video. Hãy thay bằng link bài đăng.")
    ghi_chu_douyin = ""
    try:
        return ytdlp_tai(url, outdir, name, ffmpeg_dir), "yt-dlp"
    except Exception as e:
        loi_goc = e

    # Douyin: link rút gọn chuyển về iesdouyin.com/share/video/<ID> mà yt-dlp không nhận
    m = re.search(r"iesdouyin\.com/share/video/(\d+)", str(loi_goc) + " " + url)
    if m:
        url = f"https://www.douyin.com/video/{m.group(1)}"
        don_rac(outdir, name)
        try:
            return ytdlp_tai(url, outdir, name, ffmpeg_dir), "yt-dlp (đổi link Douyin)"
        except Exception as e:
            loi_goc = e

    if "douyin" in url.lower():
        # Douyin: CHỈ lấy theo đúng mã video. Không dò bừa trong trang, vì trang Douyin phát lẫn
        # video gợi ý, clip giao diện và các đoạn tách riêng hình/tiếng -> rất dễ tải nhầm.
        ma = (m.group(1) if m else None) or douyin_id(item["link"])
        if not ma:
            m2 = re.search(r"\[Douyin\]\s*(\d+)", str(loi_goc))
            ma = m2.group(1) if m2 else None
        if not ma:
            raise RuntimeError(f"Không đọc được mã video từ link Douyin này. Lỗi gốc: {loi_goc}")
        don_rac(outdir, name)
        try:
            return tai_douyin(ma, outdir, name, ffmpeg_dir)
        except Exception as e:
            don_rac(outdir, name)
            raise RuntimeError(f"Douyin (mã {ma}): {e}. Thử chạy lại sau ít phút; nếu vẫn lỗi, mở link trên "
                               f"trình duyệt xem video còn không.")

    msg = str(loi_goc).lower()
    if "unusual" in msg and "extension" in msg:
        don_rac(outdir, name)
        try:
            return ytdlp_tai(url, outdir, name, ffmpeg_dir, ep_mp4=True), "yt-dlp (đuôi file lạ)"
        except Exception as e:
            loi_goc = e
            msg = str(e).lower()

    if re.search(r"ssl|403|cloudflare|impersonat|forbidden", msg):
        don_rac(outdir, name)
        try:
            return ytdlp_tai(url, outdir, name, ffmpeg_dir, impersonate=True), "yt-dlp (giả lập Chrome)"
        except Exception as e:
            ghi_chu_douyin += " [Giả lập Chrome cũng lỗi: " + re.sub(r"\s+", " ", str(e))[:140] + "]"

    if (TIKTOK_TRUNG_GIAN and re.search(r"tiktok\.com", url, re.I)
            and re.search(r"log in|login|comfortable|cookies|permission|private|ip address is blocked", msg)):
        print("    TikTok bắt đăng nhập -> nhờ dịch vụ trung gian tikwm.com lấy hộ (chỉ gửi địa chỉ video)...")
        don_rac(outdir, name)
        try:
            return tai_tiktok_trung_gian(url, outdir, name, ffmpeg_dir)
        except Exception as e:
            don_rac(outdir, name)
            ly_do = " ".join(str(e).split())[:140]
            raise RuntimeError(f"{loi_goc} [Dịch vụ trung gian tikwm.com cũng không lấy được: {ly_do}]")

    if (TIKTOK_TRUNG_GIAN and re.search(r"(?:twitter|x)\.com/[^/]+/status", url, re.I)
            and re.search(r"no video could be found|log in|login|cookies|nsfw|sensitive|age.restricted|authentication", msg)):
        print("    X không cho khách xem video này -> hỏi dịch vụ trung gian FxTwitter (chỉ gửi mã tweet)...")
        don_rac(outdir, name)
        try:
            return tai_x_trung_gian(url, outdir, name, ffmpeg_dir)
        except Exception as e:
            don_rac(outdir, name)
            ly_do = " ".join(str(e).split())[:140]
            raise RuntimeError(f"{loi_goc} [Dịch vụ trung gian FxTwitter cũng không lấy được: {ly_do}]")

    if NEN_TANG_LON.search(url):
        raise loi_goc

    print("    yt-dlp không lấy được -> đang dò video nhúng trong trang...")
    can_cookie = "cookie" in msg
    da_thu = set()
    loi_nguon = []
    cac_cach = [("dò HTML", do_html), ("dò trình duyệt", do_trinh_duyet)]
    if re.search(r"reddit\.com/r/[^/]+/comments/", url, re.I):
        # Reddit: chỉ lấy video nằm trong đúng bài đăng; trang Reddit còn phát video của bài gợi ý
        cac_cach = [("Reddit", do_reddit)]
    for ten_cach, ham in cac_cach:
        try:
            ung_vien = ham(url)
        except Exception as e:
            print(f"    ({ten_cach} lỗi: {str(e)[:120]})")
            ung_vien = []
        if ten_cach == "dò trình duyệt" and can_cookie and COOKIE_AUTO:
            # Trang chỉ cần cookie mới (không cần đăng nhập): dùng cookie trình duyệt ẩn vừa nhận
            don_rac(outdir, name)
            try:
                f = ytdlp_tai(url, outdir, name, ffmpeg_dir, cookiefile=COOKIE_AUTO)
                return f, "yt-dlp + cookie tự lấy bằng trình duyệt"
            except Exception:
                pass
        ung_vien = [u for u in ung_vien if u not in da_thu][:8]
        if ung_vien:
            print(f"    {ten_cach}: tìm thấy {len(ung_vien)} nguồn, đang thử...")
        for u in ung_vien:
            da_thu.add(u)
            don_rac(outdir, name)
            # Cùng một file xuất hiện ở nhiều trang khác nhau -> là file giao diện, không phải video bài đăng
            khoa = u.split("?")[0]
            if not HOST_NHUNG.search(u) and NGUON_DA_DUNG.get(khoa, url) != url:
                print(f"    (bỏ qua nguồn dùng chung giữa nhiều trang: {khoa[:80]})")
                continue
            try:
                ep = not re.search(r"\.(m3u8|mpd)(\?|$)", u, re.I) and not HOST_NHUNG.search(u)
                # Bài X/Instagram/TikTok... nhúng trong báo: tải như link gốc, KHÔNG gửi kèm địa chỉ trang báo
                # (máy chủ video của X trả lỗi 403 khi thấy Referer lạ)
                la_bai_nhung = bool(NEN_TANG_LON.search(u.split("?")[0]))
                f = ytdlp_tai(u, outdir, name, ffmpeg_dir, referer=None if la_bai_nhung else url, ep_mp4=ep,
                              cookiefile=None if la_bai_nhung else COOKIE_AUTO)
                if f.stat().st_size < KICH_THUOC_TOI_THIEU:
                    f.unlink(missing_ok=True)
                    continue
                kt = kiem_tra_luong(f, ffmpeg_dir)
                if kt and not kt[0]:
                    # chỉ có tiếng, không có hình -> không phải video của bài
                    f.unlink(missing_ok=True)
                    if len(loi_nguon) < 3:
                        loi_nguon.append(u.split("/")[2] + ": nguồn chỉ có tiếng, không có hình")
                    continue
                NGUON_DA_DUNG[khoa] = url
                don_rac(outdir, name)
                canh_bao = GHI_CHU_NGUON.get(u, "")
                if kt and not kt[1] and not canh_bao:
                    canh_bao = "CẦN KIỂM TRA: file không có tiếng"
                return f, f"{ten_cach}{' (' + canh_bao + ')' if canh_bao else ''}: {u}"
            except Exception as e:
                if len(loi_nguon) < 3:
                    loi_nguon.append(u.split("/")[2] + ": " + re.sub(r"\x1b\[[0-9;]*m|ERROR: |\s+", " ", str(e)).strip()[:70])
                continue
    don_rac(outdir, name)
    td = ("" if TRANG_THAI_TD == "ok" else f" [Dò bằng trình duyệt KHÔNG chạy: {TRANG_THAI_TD}]") + ghi_chu_douyin
    if da_thu:
        raise RuntimeError(f"Tìm thấy {len(da_thu)} nguồn video trong trang nhưng đều không tải được "
                           f"(có thể cần đăng nhập/bản quyền). [Nguồn đã thử: {' | '.join(loi_nguon)}]{td} "
                           f"Lỗi gốc: {loi_goc}")
    raise RuntimeError(f"Không tìm thấy video nhúng trong trang.{td} Lỗi gốc: {loi_goc}")


def rut_gon_loi(e: Exception) -> str:
    msg = re.sub(r"\x1b\[[0-9;]*m", "", str(e)).replace("ERROR: ", "").strip()
    low = msg.lower()
    goi_y = ""
    if "407" in low and "proxy" in low:
        goi_y = " -> PROXY từ chối: sai tài khoản/mật khẩu proxy, hoặc proxy hết hạn (kiểm tra ở Cài đặt > Proxy)"
    elif "unable to connect to proxy" in low or "tunnel connection failed" in low or "connect tunnel failed" in low:
        goi_y = " -> Không đi qua được PROXY: proxy đang tắt hoặc không vào được trang này (bấm 'Kiểm tra proxy' ở Cài đặt)"
    elif "failed to resolve" in low or "getaddrinfo" in low:
        goi_y = " -> MẠNG đang chặn tên miền này: đổi DNS sang 1.1.1.1 / 8.8.8.8 hoặc bật VPN (xem HUONG_DAN.txt mục 8)"
    elif "ip address is blocked" in low or "not available in your country" in low or "geo" in low:
        goi_y = " -> Video chặn theo quốc gia/IP: cần VPN"
    elif any(k in low for k in ("login", "log in", "cookies", "private", "rate-limit", "sign in", "logged-in",
                                "authentication", "empty media response", "cannot parse data", "no video formats")):
        goi_y = " -> Cần đăng nhập: thêm cookie của trang này (xem HUONG_DAN.txt mục 4)"
    elif "no video could be found in this tweet" in low:
        goi_y = " -> Tweet không có video trực tiếp (có thể video ở tweet được trích dẫn) hoặc cần cookie X"
    elif "ssl" in low or "eof occurred" in low:
        goi_y = " -> Kết nối bị ngắt (mạng/tường lửa): thử lại, đổi DNS hoặc bật VPN"
    elif "unsupported url" in low:
        goi_y = " -> Trang không được hỗ trợ hoặc không có video"
    elif "unavailable" in low or "removed" in low or "not available" in low:
        goi_y = " -> Video đã bị xóa/ẩn hoặc bị chặn theo khu vực"
    return (msg[:700] + goi_y).replace("\n", " ")


def chon_file_bang_hop_thoai():
    try:
        import tkinter as tk
        from tkinter import filedialog

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        p = filedialog.askopenfilename(
            title="Chọn file Excel danh sách video",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả", "*.*")],
        )
        root.destroy()
        return p or None
    except Exception:
        return None


def main():
    if len(sys.argv) >= 2:
        excel = Path(sys.argv[1].strip('"'))
    else:
        p = chon_file_bang_hop_thoai()
        if not p:
            p = input("Kéo thả file Excel vào đây rồi nhấn Enter: ").strip().strip('"')
        excel = Path(p)
    if not excel.exists():
        print(f"Không tìm thấy file: {excel}")
        return 1
    excel = excel.resolve()
    if excel.suffix.lower() == ".xls":
        print("File .xls (Excel cũ) không đọc được. Hãy mở và 'Lưu thành' .xlsx rồi chạy lại.")
        return 1

    outdir = Path(sys.argv[2]) if len(sys.argv) >= 3 else excel.parent / (excel.stem + "_video")
    outdir.mkdir(parents=True, exist_ok=True)
    outdir = outdir.resolve()
    # Chạy trong thư mục video: tránh lỗi "Permission denied ...\system32\tmp..."
    # khi tool được mở từ cửa sổ quản trị (thư mục mặc định là system32).
    os.chdir(outdir)
    tempfile.tempdir = str(outdir)

    global TRANG_THAI_TD
    dam_bao_goi()
    gop_cookie()
    TRANG_THAI_TD = kiem_tra_trinh_duyet()
    print("Dò video bằng trình duyệt ẩn: " + ("SẴN SÀNG" if TRANG_THAI_TD == "ok" else "KHÔNG CHẠY - " + TRANG_THAI_TD))
    try:
        import curl_cffi  # noqa: F401
        print("Giả lập Chrome (curl_cffi): SẴN SÀNG")
    except ImportError:
        print("Giả lập Chrome (curl_cffi): chưa cài (trang dùng Cloudflare có thể lỗi 403)")

    sheets = doc_excel(excel)
    if not sheets:
        print("Không có link nào trong file.")
        return 1

    ffmpeg_dir = chuan_bi_ffmpeg()
    if not ffmpeg_dir:
        print("! Không có ffmpeg: YouTube có thể chỉ tải được tối đa 720p.")

    # Nhiều sheet -> mỗi sheet một thư mục con (vì tên file như 2.1 lặp lại giữa các sheet)
    nhieu_sheet = len(sheets) > 1
    tong = sum(len(it) for _, it in sheets)
    print(f"\nTổng cộng {tong} link trong {len(sheets)} sheet. Lưu vào: {outdir}\n")
    ket_qua = []
    ok = skip = fail = 0
    dem = 0
    dung = False
    for ten_sheet, items in sheets:
        thu_muc = outdir / ten_hop_le(ten_sheet) if nhieu_sheet else outdir
        thu_muc.mkdir(exist_ok=True)
        if nhieu_sheet:
            print(f"\n########## SHEET: {ten_sheet}  ->  {thu_muc.name}\\ ##########")
        for it in items:
            dem += 1
            print(f"[{dem}/{tong}] {it['name']}  <-  {it['link']}")
            existed = da_tai(thu_muc, it["name"])
            if existed:
                print(f"    Đã có sẵn ({existed.name}), bỏ qua.")
                existed, ghi_chu = dam_bao_h264(existed, ffmpeg_dir)
                skip += 1
                ket_qua.append([ten_sheet, it["row"], it["name"], it["link"], "Đã có sẵn", existed.name, ghi_chu, ""])
                continue
            try:
                f, cach = tai_mot(it, thu_muc, ffmpeg_dir)
                print(f"    OK -> {f.name}  ({cach[:90]})")
                ok += 1
                ket_qua.append([ten_sheet, it["row"], it["name"], it["link"], "Thành công", f.name, cach, ""])
            except KeyboardInterrupt:
                print("\nĐã dừng theo yêu cầu.")
                dung = True
                break
            except Exception as e:
                loi = rut_gon_loi(e)
                print(f"    LỖI: {loi}")
                fail += 1
                ket_qua.append([ten_sheet, it["row"], it["name"], it["link"], "Lỗi", "", "", loi])
        if dung:
            break

    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    report = outdir / f"_ket_qua_{stamp}.csv"
    with open(report, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Sheet", "Hàng Excel", "Tên file", "Link", "Trạng thái", "File đã lưu", "Cách lấy", "Lỗi"])
        w.writerows(ket_qua)

    print("\n================ KẾT QUẢ ================")
    print(f"Thành công: {ok} | Đã có sẵn: {skip} | Lỗi: {fail}")
    print(f"Thư mục video: {outdir}")
    print(f"Báo cáo:       {report}")
    if fail:
        print("Chạy lại công cụ để thử tải lại các link lỗi (file đã tải sẽ được bỏ qua).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
