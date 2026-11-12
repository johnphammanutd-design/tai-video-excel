#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GIAO DIỆN cho công cụ tải video từ Excel.
Chạy:  python giao_dien.py   (hoặc nhấp đúp MO_GIAO_DIEN.bat)
Giao diện mở trong cửa sổ Edge/Chrome, chỉ chạy trên máy này (127.0.0.1).
"""
import csv
import datetime
import json
import os
import secrets
import shutil
import socket
import subprocess
import sys
import tempfile
import threading
import time
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import tai_video as T  # noqa: E402

AN_CUA_SO = getattr(subprocess, "CREATE_NO_WINDOW", 0)   # không bật cửa sổ đen khi gọi lệnh con
TOKEN = secrets.token_urlsafe(24)
LOCK = threading.RLock()


class DungTai(BaseException):
    """Người dùng bấm Dừng."""


S = {
    "items": [], "excel": "", "outdir": "", "running": False, "stopping": False, "cur": None,
    "env": {"ready": False}, "report": "", "last_poll": time.time(), "next_id": 1, "ffmpeg": None,
    "luong": 3,          # số link tải cùng lúc (mỗi nền tảng vẫn chỉ 1 link một lúc)
    "cap_nhat": None,    # thông tin bản mới trên GitHub (nếu có)
}
LOG = []
LUONG_TOI_DA = 6
NHAN = threading.local()   # .ten = nhãn ghi trước mỗi dòng nhật ký của luồng; .cur = link luồng đang tải


class Tee:
    def __init__(self, goc):
        self.goc = goc
        self.bufs = {}
        self.khoa = threading.Lock()

    def write(self, s):
        try:
            if self.goc:
                self.goc.write(s)
        except Exception:
            pass
        # Mỗi luồng một bộ đệm riêng để các dòng nhật ký của nhiều link tải cùng lúc không trộn vào nhau
        with self.khoa:
            k = threading.get_ident()
            buf = self.bufs.pop(k, "") + s
            ten = getattr(NHAN, "ten", "")
            while "\n" in buf:
                line, buf = buf.split("\n", 1)
                line = line.rstrip("\r")
                if "\r" in line:
                    line = line.split("\r")[-1]
                if line.strip():
                    LOG.append(datetime.datetime.now().strftime("%H:%M:%S  ")
                               + (f"[{ten}] {line.strip()}" if ten else line))
                    del LOG[:-800]
            if buf:
                self.bufs[k] = buf

    def flush(self):
        try:
            if self.goc:
                self.goc.flush()
        except Exception:
            pass


def thu_muc_mac_dinh():
    return str(Path.home() / "Downloads" / "TaiVideo")


def thu_muc_item(sheet):
    base = Path(S["outdir"] or thu_muc_mac_dinh())
    sheets = {i["sheet"] for i in S["items"]} | {sheet}
    return str(base / T.ten_hop_le(sheet)) if len(sheets) > 1 else str(base)


def tim_file_nghi_sai():
    """Đánh dấu file nghi tải sai: quá nhỏ, hoặc nhiều link khác nhau cho ra file giống hệt nhau."""
    import hashlib

    theo_co = {}
    for it in S["items"]:
        it["suspect"] = ""
        if it["status"] not in ("exists", "done") or not it["file"]:
            continue
        f = Path(it["dir"]) / it["file"]
        try:
            co = f.stat().st_size
        except OSError:
            continue
        if co < T.KICH_THUOC_TOI_THIEU:
            it["suspect"] = "File quá nhỏ, có thể không phải video"
        theo_co.setdefault(co, []).append((it, f))
    for co, nhom in theo_co.items():
        if len(nhom) < 2:
            continue
        theo_bam = {}
        for it, f in nhom:
            try:
                h = hashlib.md5()
                with open(f, "rb") as fh:
                    for khoi in iter(lambda: fh.read(1 << 20), b""):
                        h.update(khoi)
                theo_bam.setdefault(h.hexdigest(), []).append(it)
            except OSError:
                pass
        for ds in theo_bam.values():
            if len({i["link"] for i in ds}) > 1:
                for i in ds:
                    i["suspect"] = f"Giống hệt {len(ds) - 1} file khác dù link khác nhau: nhiều khả năng tải sai"


def sua_file_nghi_sai():
    with LOCK:
        if S["running"]:
            raise ValueError("Đang tải, hãy dừng trước.")
        n = 0
        for it in S["items"]:
            if it.get("suspect") and it["file"]:
                try:
                    (Path(it["dir"]) / it["file"]).unlink()
                except OSError:
                    continue
                it.update(status="pending", file="", method="", suspect="", pct=0)
                n += 1
    print(f"Đã xóa {n} file nghi tải sai, sẽ tải lại.")
    if n:
        bat_dau("all")


def nap_file_tai_len(ten, du_lieu):
    """Người dùng kéo thả file Excel vào cửa sổ: lưu bản sao rồi mở."""
    ten = T.ten_hop_le(Path(ten).name) or "danh_sach.xlsx"
    if not ten.lower().endswith((".xlsx", ".xlsm")):
        raise ValueError("Chỉ nhận file Excel .xlsx hoặc .xlsm.")
    d = T.TAM_DIR / "tai_len"
    d.mkdir(parents=True, exist_ok=True)
    (d / ten).write_bytes(du_lieu)
    nap_excel(str(d / ten))
    with LOCK:
        # Không biết thư mục gốc của file -> lưu video vào Downloads\TaiVideo\<tên file>_video
        S["excel"] = ten
        S["outdir"] = str(Path(thu_muc_mac_dinh()) / (Path(ten).stem + "_video"))
        cap_nhat_thu_muc()


def cap_nhat_thu_muc():
    """Tính lại thư mục lưu của từng link và đánh dấu link đã có file."""
    try:
        _cap_nhat_thu_muc()
    finally:
        tim_file_nghi_sai()


def _cap_nhat_thu_muc():
    for it in S["items"]:
        it["dir"] = thu_muc_item(it["sheet"])
        if it["status"] in ("pending", "exists", "error"):
            f = T.da_tai(Path(it["dir"]), it["name"]) if Path(it["dir"]).is_dir() else None
            if f:
                it.update(status="exists", file=f.name, error="")
            elif it["status"] == "exists":
                it.update(status="pending", file="")


def them_item(sheet, row, name, link):
    it = {"id": S["next_id"], "sheet": sheet, "row": row, "name": name, "link": link, "status": "pending",
          "pct": 0, "method": "", "error": "", "file": "", "dir": "", "suspect": ""}
    S["next_id"] += 1
    S["items"].append(it)
    return it


def nap_excel(path):
    p = Path(path.strip().strip('"'))
    if not p.exists():
        raise ValueError(f"Không tìm thấy file: {p}")
    if p.suffix.lower() == ".xls":
        raise ValueError("File .xls (Excel cũ) không đọc được. Hãy lưu lại thành .xlsx.")
    sheets = T.doc_excel(p)
    if not sheets:
        raise ValueError("Không tìm thấy link nào. File cần có cột 'Link video' và 'Đặt tên file tải về'.")
    with LOCK:
        if S["running"]:
            raise ValueError("Đang tải, hãy dừng trước khi mở file khác.")
        S["items"] = [i for i in S["items"] if i["sheet"] == "Link lẻ"]
        S["excel"] = str(p.resolve())
        S["outdir"] = str(p.resolve().parent / (p.stem + "_video"))
        for ten, items in sheets:
            for it in items:
                them_item(ten, it["row"], it["name"], it["link"])
        cap_nhat_thu_muc()
    print(f"Đã mở {p.name}: {sum(len(i) for _, i in sheets)} link trong {len(sheets)} sheet.")


def them_link(text, ten=""):
    links = [x.strip() for x in text.replace(",", "\n").split() if x.strip()]
    links = [x if x.lower().startswith("http") else "https://" + x for x in links if "." in x]
    if not links:
        raise ValueError("Chưa có link hợp lệ.")
    with LOCK:
        da_co = {i["name"] for i in S["items"] if i["sheet"] == "Link lẻ"}
        n = len(da_co)
        for k, u in enumerate(links):
            name = T.ten_hop_le(ten) if ten and len(links) == 1 else ""
            if ten and len(links) > 1:
                name = T.ten_hop_le(f"{ten}_{k + 1}")
            while not name or name in da_co:
                n += 1
                name = f"video_{n}"
            da_co.add(name)
            them_item("Link lẻ", 0, name, u)
        if not S["outdir"]:
            S["outdir"] = thu_muc_mac_dinh()
        cap_nhat_thu_muc()
    return len(links)


def hook(d):
    if S["stopping"]:
        raise DungTai()
    cur = getattr(NHAN, "cur", None)   # link mà CHÍNH luồng này đang tải
    if cur and d.get("status") == "downloading":
        tong = d.get("total_bytes") or d.get("total_bytes_estimate")
        if tong:
            cur["pct"] = max(cur["pct"], min(99, int(d.get("downloaded_bytes", 0) * 100 / tong)))


def ghi_bao_cao(ids):
    rows = [i for i in S["items"] if i["id"] in ids and i["status"] in ("done", "exists", "error")]
    if not rows or not S["outdir"]:
        return
    nhan = {"done": "Thành công", "exists": "Đã có sẵn", "error": "Lỗi"}
    out = Path(S["outdir"])
    out.mkdir(parents=True, exist_ok=True)
    f = out / ("_ket_qua_" + datetime.datetime.now().strftime("%Y%m%d_%H%M%S") + ".csv")
    with open(f, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["Sheet", "Hàng Excel", "Tên file", "Link", "Trạng thái", "File đã lưu", "Cách lấy", "Lỗi"])
        for i in rows:
            w.writerow([i["sheet"], i["row"], i["name"], i["link"], nhan[i["status"]], i["file"], i["method"], i["error"]])
    S["report"] = str(f)
    print(f"Báo cáo: {f}")


def tho(ids):
    try:
        ENV_THREAD.join()
        T.PROGRESS_HOOK = hook
        T.gop_cookie()
        if S["ffmpeg"] is None:
            S["ffmpeg"] = T.chuan_bi_ffmpeg() or ""
        ds = [i for i in S["items"] if i["id"] in ids]
        # Thư mục làm việc và thư mục tạm là của CẢ chương trình nên đặt một lần cho cả lượt
        # (trước đây đổi theo từng link; nhiều luồng cùng đổi sẽ giẫm lên nhau).
        goc = Path(S["outdir"] or thu_muc_mac_dinh())
        goc.mkdir(parents=True, exist_ok=True)
        os.chdir(goc)
        tempfile.tempdir = str(goc)

        hang = list(ds)            # các link chưa ai nhận
        ban = set()                # các nền tảng đang có luồng tải
        cv = threading.Condition()
        dem = [0]

        def nhan_viec():
            """Lấy link kế tiếp mà nền tảng của nó đang rảnh. None = hết việc hoặc người dùng bấm Dừng."""
            with cv:
                while not S["stopping"]:
                    if not hang:
                        return None
                    for it in hang:
                        nt = T.nen_tang(it["link"])
                        if nt not in ban:
                            hang.remove(it)
                            ban.add(nt)
                            dem[0] += 1
                            return it, nt, dem[0]
                    cv.wait(0.5)   # mọi link còn lại đều thuộc nền tảng đang bận -> chờ
                return None

        def cong_nhan():
            while True:
                viec = nhan_viec()
                if not viec:
                    return
                it, nt, so = viec
                try:
                    tai_mot_link(it, so, len(ds))
                except BaseException as e:   # không để một luồng chết làm treo cả lượt
                    it.update(status="error", error=f"Lỗi không mong muốn: {e}", pct=0)
                finally:
                    NHAN.ten, NHAN.cur = "", None
                    with cv:
                        ban.discard(nt)
                        cv.notify_all()

        n = max(1, min(int(S["luong"] or 1), LUONG_TOI_DA, len(ds)))
        if n > 1:
            print(f"Tải {n} link cùng lúc (mỗi nền tảng một link một lúc).")
        cac_luong = [threading.Thread(target=cong_nhan, daemon=True) for _ in range(n)]
        for t in cac_luong:
            t.start()
        for t in cac_luong:
            t.join()
        ghi_bao_cao(ids)
        tim_file_nghi_sai()
    except Exception as e:
        print(f"Lỗi không mong muốn: {e}")
    finally:
        for it in S["items"]:
            if it["status"] == "running":
                it["status"] = "pending"
        os.chdir(HERE)
        T.xoa_ban_sao_cookie()
        da_dung = S["stopping"]
        S.update(running=False, stopping=False, cur=None)
        print("Đã dừng." if da_dung else "Hoàn tất lượt tải.")


def tai_mot_link(it, so, tong):
    """Tải một link (chạy trong một luồng công nhân)."""
    d = Path(it["dir"])
    d.mkdir(parents=True, exist_ok=True)
    co = T.da_tai(d, it["name"])
    if co:
        it.update(status="exists", file=co.name, error="", pct=100)
        return
    print(f"[{so}/{tong}] {it['sheet']} / {it['name']}  <-  {it['link']}")
    NHAN.ten = f"{it['sheet']} / {it['name']}"
    NHAN.cur = it
    it.update(status="running", pct=0, error="", method="")
    try:
        f, cach = T.tai_mot({"link": it["link"], "name": it["name"]}, d, S["ffmpeg"] or None)
        it.update(status="done", file=f.name, method=cach, pct=100)
        print(f"OK -> {f.name}  ({cach[:90]})")
    except DungTai:
        T.don_rac(d, it["name"])
        it.update(status="pending", pct=0)
    except Exception as e:
        loi = T.rut_gon_loi(e)
        it.update(status="error", error=loi, pct=0)
        print(f"LỖI: {loi}")


# ======================= CÀI ĐẶT LƯU LẠI GIỮA CÁC LẦN MỞ =======================
def _file_cai_dat():
    return T.TAM_DIR / "cai_dat.json"


def nap_cai_dat():
    try:
        d = json.loads(_file_cai_dat().read_text(encoding="utf-8"))
        S["luong"] = max(1, min(LUONG_TOI_DA, int(d.get("luong") or 3)))
        T.MAX_HEIGHT = int(d.get("max_height") or 0) or None
    except Exception:
        pass


def luu_cai_dat():
    try:
        T.TAM_DIR.mkdir(exist_ok=True)
        _file_cai_dat().write_text(json.dumps({"luong": S["luong"], "max_height": T.MAX_HEIGHT or 0}),
                                   encoding="utf-8")
    except Exception:
        pass


# ======================= TỰ KIỂM TRA VÀ CẬP NHẬT BẢN MỚI =======================
# Nguồn: kho GitHub công khai của tool. File phien_ban.json (nhánh main) ghi số phiên bản mới nhất,
# thẻ (tag) chứa bản đó và mã băm SHA-256 của từng file. Tool CHỈ báo có bản mới; người dùng bấm mới cập nhật.
KHO_CAP_NHAT = "johnphammanutd-design/tai-video-excel"
FILE_DUOC_CAP_NHAT = {"tai_video.py", "giao_dien.py", "giao_dien.html", "HUONG_DAN.txt"}


def _tai_tu_kho(nhanh, ten, toi_da=5_000_000):
    import urllib.request

    thu = os.environ.get("TAIVIDEO_UPDATE_URL")   # chỉ để chạy thử bằng máy chủ trên chính máy này
    if thu:
        if not thu.startswith(("http://127.0.0.1", "http://localhost")):
            raise ValueError("TAIVIDEO_UPDATE_URL chỉ dùng được với máy chủ trên máy này.")
        nguon = [f"{thu.rstrip('/')}/{nhanh}/"]
    else:
        nguon = [f"https://raw.githubusercontent.com/{KHO_CAP_NHAT}/{nhanh}/",
                 f"https://cdn.jsdelivr.net/gh/{KHO_CAP_NHAT}@{nhanh}/"]
    loi = None
    for goc in nguon:
        try:
            req = urllib.request.Request(goc + ten, headers={"User-Agent": "TaiVideoExcel"})
            with urllib.request.urlopen(req, timeout=20) as r:
                du_lieu = r.read(toi_da + 1)
            if len(du_lieu) > toi_da:
                raise ValueError("file quá lớn")
            return du_lieu
        except Exception as e:
            loi = e
    raise loi


def kiem_tra_ban_moi():
    """Hỏi kho GitHub xem có bản mới hơn không. Không làm gì khác ngoài ghi nhận để giao diện báo."""
    if os.environ.get("TAIVIDEO_NO_UPDATE") or (HERE / ".git").exists():
        return   # thư mục của người phát triển: không tự ghi đè code đang sửa
    try:
        d = json.loads(_tai_tu_kho("main", "phien_ban.json", 100_000).decode("utf-8"))
        moi, the, files = str(d["phien_ban"]), str(d["the"]), dict(d["files"])
        if int(moi) <= int(T.PHIEN_BAN):
            return
        import re
        if not re.fullmatch(r"v\d+", the) or not files or not set(files) <= FILE_DUOC_CAP_NHAT \
                or not all(re.fullmatch(r"[0-9a-f]{64}", str(b)) for b in files.values()):
            print("Thông tin bản mới trên GitHub không hợp lệ, bỏ qua.")
            return
        S["cap_nhat"] = {"moi": moi, "the": the, "files": files, "ghi_chu": str(d.get("ghi_chu") or "")[:400]}
        print(f"Có bản mới: phiên bản {moi} (đang dùng {T.PHIEN_BAN}). Bấm 'Cập nhật' ở góc dưới bên trái.")
    except Exception as e:
        print(f"Không kiểm tra được bản mới ({str(e)[:80]}). Vẫn dùng bình thường.")


def cap_nhat_tool():
    """Tải bản mới, kiểm tra từng file rồi mới thay; giữ bản cũ để quay lại; mở lại tool."""
    import hashlib

    cn = S.get("cap_nhat")
    if not cn:
        raise ValueError("Không có bản mới.")
    with LOCK:
        if S["running"]:
            raise ValueError("Đang tải, hãy dừng hoặc chờ xong rồi cập nhật.")
    moi = {}
    for ten, bam in cn["files"].items():
        try:
            du_lieu = _tai_tu_kho(cn["the"], ten)
        except Exception as e:
            raise ValueError(f"Không tải được bản mới ({ten}): {str(e)[:80]}. Thử lại sau.")
        if hashlib.sha256(du_lieu).hexdigest() != bam:
            raise ValueError(f"File {ten} tải về không khớp mã kiểm tra. Đã hủy cập nhật, tool giữ nguyên bản cũ.")
        if ten.endswith(".py"):
            try:
                compile(du_lieu.decode("utf-8"), ten, "exec")
            except Exception:
                raise ValueError(f"File {ten} của bản mới bị lỗi. Đã hủy cập nhật, tool giữ nguyên bản cũ.")
        moi[ten] = du_lieu
    # Giữ bản đang dùng để quay lại khi cần
    cu = T.TAM_DIR / f"ban_cu_{T.PHIEN_BAN}"
    cu.mkdir(parents=True, exist_ok=True)
    for ten in moi:
        if (HERE / ten).exists():
            shutil.copy2(HERE / ten, cu / ten)
    da_thay = []
    try:
        for ten, du_lieu in moi.items():
            tam = HERE / (ten + ".moi")
            tam.write_bytes(du_lieu)
            os.replace(tam, HERE / ten)
            da_thay.append(ten)
    except Exception as e:
        for ten in da_thay:   # thay dở dang -> trả lại bản cũ
            try:
                shutil.copy2(cu / ten, HERE / ten)
            except Exception:
                pass
        raise ValueError(f"Không ghi được file vào thư mục cài đặt ({str(e)[:80]}). Tool giữ nguyên bản cũ.")
    print(f"Đã cập nhật lên phiên bản {cn['moi']}. Bản cũ lưu ở {cu}. Đang mở lại tool...")
    env = {k: v for k, v in os.environ.items() if k != "TAIVIDEO_PORT"}
    lenh = [sys.executable, str(HERE / "giao_dien.py")]
    if S["excel"] and Path(S["excel"]).exists():
        lenh.append(S["excel"])
    subprocess.Popen(lenh, cwd=str(HERE), env=env, creationflags=AN_CUA_SO)
    S["thoat"] = time.time()


def bat_dau(che_do, ids=None):
    with LOCK:
        if S["running"]:
            return
        cap_nhat_thu_muc()
        if ids:
            chon = [i for i in S["items"] if i["id"] in ids and i["status"] != "running"]
        elif che_do == "failed":
            chon = [i for i in S["items"] if i["status"] == "error"]
        else:
            chon = [i for i in S["items"] if i["status"] in ("pending", "error")]
        if not chon:
            raise ValueError("Không có link nào cần tải.")
        for i in chon:
            if i["status"] == "exists" and ids:
                continue
            i.update(status="pending", error="", pct=0)
        S.update(running=True, stopping=False)
        threading.Thread(target=tho, args=({i["id"] for i in chon},), daemon=True).start()


def cap_nhat_ytdlp():
    """Cập nhật yt-dlp tối đa 1 lần/ngày (các trang đổi cách chặn tải liên tục)."""
    moc = T.TAM_DIR / "cap_nhat_ytdlp.txt"
    hom_nay = datetime.date.today().isoformat()
    try:
        if moc.read_text().strip() == hom_nay:
            return
    except Exception:
        pass
    print("Đang cập nhật yt-dlp...")
    try:
        r = subprocess.run([sys.executable, "-m", "pip", "install", "-q", "-U", "--disable-pip-version-check",
                            "yt-dlp[default,curl-cffi]"], capture_output=True, text=True, timeout=600,
                           creationflags=AN_CUA_SO)
        if r.returncode == 0:
            T.TAM_DIR.mkdir(exist_ok=True)
            moc.write_text(hom_nay)
            print("    yt-dlp đã là bản mới nhất.")
        else:
            print("    Không cập nhật được yt-dlp (mất mạng?). Dùng bản đang có.")
    except Exception as e:
        print(f"    Không cập nhật được yt-dlp: {e}")


def kiem_tra_moi_truong():
    try:
        if not os.environ.get("TAIVIDEO_NO_UPDATE"):
            cap_nhat_ytdlp()
        T.dam_bao_goi()
        T.TRANG_THAI_TD = T.kiem_tra_trinh_duyet()
    except Exception as e:
        print(f"Kiểm tra môi trường lỗi: {e}")
    env = {"ready": True, "python": sys.version.split()[0], "browser": T.TRANG_THAI_TD}
    try:
        import yt_dlp
        env["ytdlp"] = yt_dlp.version.__version__
    except Exception:
        env["ytdlp"] = ""
    try:
        import curl_cffi  # noqa: F401
        env["curl"] = True
    except Exception:
        env["curl"] = False
    S["ffmpeg"] = T.chuan_bi_ffmpeg() or ""
    env["ffmpeg"] = bool(S["ffmpeg"])
    S["env"] = env
    kiem_tra_ban_moi()
    print("Sẵn sàng." if env["ytdlp"] else "THIẾU yt-dlp: không tải được. Kiểm tra kết nối mạng rồi mở lại.")


def chon_bang_hop_thoai(loai):
    if os.name == "nt":
        # Hộp thoại gốc của Windows qua PowerShell (bản cài đặt không kèm tkinter)
        if loai == "excel":
            hop = ("$d=New-Object System.Windows.Forms.OpenFileDialog;$d.Title='Chon file Excel danh sach video';"
                   "$d.Filter='Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Tat ca (*.*)|*.*';"
                   "if($d.ShowDialog($f) -eq 'OK'){$p=$d.FileName}")
        else:
            hop = ("$d=New-Object System.Windows.Forms.FolderBrowserDialog;$d.Description='Chon thu muc luu video';"
                   "if($d.ShowDialog($f) -eq 'OK'){$p=$d.SelectedPath}")
        # Cửa sổ chủ vô hình, luôn nằm trên cùng và PHẢI được hiện (Show) thì hộp thoại mới nổi lên trước
        ps = ("Add-Type -AssemblyName System.Windows.Forms;$p='';"
              "$f=New-Object System.Windows.Forms.Form;$f.TopMost=$true;$f.ShowInTaskbar=$false;"
              "$f.FormBorderStyle='None';$f.StartPosition='CenterScreen';$f.Width=1;$f.Height=1;$f.Opacity=0;"
              "$f.Show();$f.Activate();[void]$f.Focus();" + hop +
              ";$f.Dispose();[Console]::OutputEncoding=[Text.Encoding]::UTF8;[Console]::Out.Write($p)")
        try:
            r = subprocess.run(["powershell", "-NoProfile", "-STA", "-ExecutionPolicy", "Bypass", "-Command", ps],
                               capture_output=True, timeout=900, creationflags=AN_CUA_SO)
            if r.returncode == 0:
                return r.stdout.decode("utf-8", "ignore").strip().lstrip("\ufeff")
        except Exception:
            pass
    if loai == "excel":
        lenh = ("filedialog.askopenfilename(title='Chọn file Excel danh sách video',"
                "filetypes=[('Excel','*.xlsx *.xlsm'),('Tất cả','*.*')])")
    else:
        lenh = "filedialog.askdirectory(title='Chọn thư mục lưu video')"
    code = ("import sys,tkinter as tk\nfrom tkinter import filedialog\nr=tk.Tk();r.withdraw();r.attributes('-topmost',True)\n"
            f"p={lenh}\nsys.stdout.buffer.write((p or '').encode('utf-8'))")
    r = subprocess.run([sys.executable, "-c", code], capture_output=True, timeout=900, creationflags=AN_CUA_SO)
    if r.returncode != 0:
        raise ValueError("Không mở được hộp thoại chọn file. Hãy kéo thả file Excel vào biểu tượng của tool.")
    return r.stdout.decode("utf-8", "ignore").strip()


def mo(path):
    p = Path(path)
    p_mo = p if p.exists() else p.parent
    if not p_mo.exists():
        p_mo.mkdir(parents=True, exist_ok=True)
    if os.name == "nt":
        os.startfile(str(p_mo))  # noqa
    else:
        subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(p_mo)])


def ds_cookie():
    out = []
    files = []
    if (T.SCRIPT_DIR / "cookies.txt").exists():
        files.append(T.SCRIPT_DIR / "cookies.txt")
    if (T.SCRIPT_DIR / "cookies").is_dir():
        files += sorted((T.SCRIPT_DIR / "cookies").glob("*.txt"))
    for f in files:
        ds = T._doc_dong_cookie(f)
        mien = sorted({x.split("\t")[0].replace("#HttpOnly_", "").lstrip(".") for x in ds})
        goc = sorted({".".join(m.split(".")[-2:]) for m in mien})
        out.append({"name": f.name, "lines": len(ds), "sites": goc[:8]})
    return out


def trang_thai(log_tu=0):
    with LOCK:
        dem = {}
        for i in S["items"]:
            dem[i["status"]] = dem.get(i["status"], 0) + 1
        return {
            "items": S["items"], "excel": S["excel"], "outdir": S["outdir"] or thu_muc_mac_dinh(),
            "running": S["running"], "stopping": S["stopping"], "env": S["env"], "report": S["report"],
            "count": dem, "cookies": ds_cookie(), "max_height": T.MAX_HEIGHT or 0,
            "log": LOG[-300:], "cur": S["cur"]["id"] if S["cur"] else None, "version": T.PHIEN_BAN,
            "suspect": sum(1 for i in S["items"] if i.get("suspect")),
            "threads": S["luong"], "threads_max": LUONG_TOI_DA,
            "update": ({"version": S["cap_nhat"]["moi"], "note": S["cap_nhat"]["ghi_chu"]} if S["cap_nhat"] else None),
            "restarting": bool(S.get("thoat")),
        }


class H(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _gui(self, code, body, ctype="application/json; charset=utf-8"):
        if not isinstance(body, bytes):
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _hop_le(self):
        host = (self.headers.get("Host") or "").split(":")[0]
        return host in ("127.0.0.1", "localhost") and self.headers.get("X-Token") == TOKEN

    def do_GET(self):
        if self.path.split("?")[0] == "/":
            html = (HERE / "giao_dien.html").read_text(encoding="utf-8").replace("__TOKEN__", TOKEN)
            return self._gui(200, html.encode("utf-8"), "text/html; charset=utf-8")
        if self.path.startswith("/favicon.ico") and (HERE / "icon.ico").exists():
            return self._gui(200, (HERE / "icon.ico").read_bytes(), "image/x-icon")
        if self.path.startswith("/api/bye"):
            # Cửa sổ giao diện vừa đóng (hoặc tải lại): thoát nếu sau 4 giây không thấy mở lại
            if self.path.endswith("t=" + TOKEN):
                S["bye"] = time.time()
            return self._gui(200, {"ok": True})
        if self.path.startswith("/api/state"):
            if not self._hop_le():
                return self._gui(403, {"error": "forbidden"})
            S["last_poll"] = time.time()
            return self._gui(200, trang_thai())
        self._gui(404, {"error": "not found"})

    def do_POST(self):
        if not self._hop_le():
            return self._gui(403, {"error": "forbidden"})
        n = int(self.headers.get("Content-Length") or 0)
        than = self.rfile.read(n) if n else b""
        if self.path.startswith("/api/upload"):
            from urllib.parse import urlsplit, parse_qs
            try:
                ten = parse_qs(urlsplit(self.path).query).get("name", ["danh_sach.xlsx"])[0]
                nap_file_tai_len(ten, than)
                return self._gui(200, {"ok": True})
            except Exception as e:
                return self._gui(200, {"error": str(e)})
        try:
            d = json.loads(than or b"{}")
        except Exception:
            d = {}
        a = self.path
        try:
            if a == "/api/pick_excel":
                p = chon_bang_hop_thoai("excel")
                if p:
                    nap_excel(p)
            elif a == "/api/load":
                nap_excel(d.get("path", ""))
            elif a == "/api/pick_outdir":
                p = chon_bang_hop_thoai("dir")
                if p:
                    with LOCK:
                        if S["running"]:
                            raise ValueError("Đang tải, hãy dừng trước khi đổi thư mục.")
                        S["outdir"] = p
                        cap_nhat_thu_muc()
            elif a == "/api/add_links":
                them_link(d.get("links", ""), d.get("name", ""))
                if d.get("start") and not S["running"]:
                    bat_dau("all")
            elif a == "/api/start":
                bat_dau(d.get("mode", "all"), set(d["ids"]) if d.get("ids") else None)
            elif a == "/api/fix_suspect":
                sua_file_nghi_sai()
            elif a == "/api/stop":
                if S["running"]:
                    S["stopping"] = True
                    print("Đang dừng sau bước hiện tại...")
            elif a == "/api/clear":
                with LOCK:
                    if S["running"]:
                        raise ValueError("Đang tải, hãy dừng trước.")
                    S.update(items=[], excel="", report="")
            elif a == "/api/open":
                w = d.get("what")
                if w == "outdir":
                    mo(S["outdir"] or thu_muc_mac_dinh())
                elif w == "cookies":
                    (T.SCRIPT_DIR / "cookies").mkdir(exist_ok=True)
                    mo(T.SCRIPT_DIR / "cookies")
                elif w == "report" and S["report"]:
                    mo(S["report"])
                elif w == "item":
                    it = next((i for i in S["items"] if i["id"] == d.get("id")), None)
                    if it:
                        mo(Path(it["dir"]) / it["file"] if it["file"] else it["dir"])
                elif w == "guide":
                    mo(HERE / "HUONG_DAN.txt")
            elif a == "/api/settings":
                if "max_height" in d:
                    T.MAX_HEIGHT = int(d.get("max_height") or 0) or None
                if "threads" in d:
                    S["luong"] = max(1, min(LUONG_TOI_DA, int(d.get("threads") or 1)))
                luu_cai_dat()
            elif a == "/api/update":
                cap_nhat_tool()
            else:
                return self._gui(404, {"error": "not found"})
            self._gui(200, {"ok": True})
        except ValueError as e:
            self._gui(200, {"error": str(e)})
        except Exception as e:
            self._gui(200, {"error": f"{type(e).__name__}: {e}"})


def mo_cua_so(url):
    if os.environ.get("TAIVIDEO_NO_BROWSER"):
        return
    ung_vien = []
    if os.name == "nt":
        for goc in (os.environ.get("ProgramFiles(x86)"), os.environ.get("ProgramFiles"), os.environ.get("LocalAppData")):
            if goc:
                ung_vien += [Path(goc) / "Microsoft/Edge/Application/msedge.exe",
                             Path(goc) / "Google/Chrome/Application/chrome.exe"]
    for exe in ung_vien:
        if exe.exists():
            try:
                subprocess.Popen([str(exe), f"--app={url}", "--window-size=1320,860",
                                  "--disable-features=Translate,msEdgeTranslate"])
                return
            except Exception:
                pass
    webbrowser.open(url)


def main():
    global ENV_THREAD
    sys.stdout = Tee(sys.stdout)
    if sys.stderr is None:   # chạy bằng pythonw (không có cửa sổ đen): ghi lỗi ra file
        try:
            T.TAM_DIR.mkdir(exist_ok=True)
            sys.stderr = open(T.TAM_DIR / "loi.log", "a", encoding="utf-8")
        except Exception:
            pass
    nap_cai_dat()
    port = int(os.environ.get("TAIVIDEO_PORT") or 0)
    if not port:
        s = socket.socket()
        s.bind(("127.0.0.1", 0))
        port = s.getsockname()[1]
        s.close()
    srv = ThreadingHTTPServer(("127.0.0.1", port), H)
    url = f"http://127.0.0.1:{port}/"
    print(f"Giao diện: {url}")
    ENV_THREAD = threading.Thread(target=kiem_tra_moi_truong, daemon=True)
    ENV_THREAD.start()
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    if len(sys.argv) > 1:
        try:
            nap_excel(sys.argv[1])
        except Exception as e:
            print(f"Không mở được file: {e}")
    mo_cua_so(url)
    try:
        while True:
            time.sleep(2)
            if S.get("thoat") and time.time() - S["thoat"] > 3:
                print("Đã cập nhật -> đóng bản cũ.")
                break
            if S["running"]:
                continue
            bye = S.get("bye", 0)
            if (bye and time.time() - bye > 4 and S["last_poll"] < bye) or time.time() - S["last_poll"] > 600:
                print("Cửa sổ giao diện đã đóng -> thoát.")
                break
    except KeyboardInterrupt:
        pass
    srv.shutdown()


if __name__ == "__main__":
    main()
