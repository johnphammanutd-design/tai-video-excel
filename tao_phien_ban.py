#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Dành cho NGƯỜI PHÁT HÀNH (không nằm trong bộ cài).
Tạo file phien_ban.json để tool trên máy mọi người biết có bản mới.

Quy trình phát hành bản N:
  1. Sửa code, tăng PHIEN_BAN trong tai_video.py và PhienBan trong dong_goi_exe/TaiVideo.iss.
  2. python tao_phien_ban.py "Ghi chú ngắn về bản này"
  3. git add -A ; git commit ; git tag vN ; git push origin main vN
Tool tải file từ thẻ vN và so mã băm SHA-256 ghi trong phien_ban.json trước khi thay.
"""
import hashlib
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
FILES = ["tai_video.py", "giao_dien.py", "giao_dien.html", "HUONG_DAN.txt"]


def main():
    m = re.search(r'^PHIEN_BAN\s*=\s*"(\d+)"', (HERE / "tai_video.py").read_text(encoding="utf-8"), re.M)
    if not m:
        print("Không đọc được PHIEN_BAN trong tai_video.py")
        return 1
    d = {
        "phien_ban": m.group(1),
        "the": "v" + m.group(1),
        "ghi_chu": sys.argv[1] if len(sys.argv) > 1 else "",
        "files": {f: hashlib.sha256((HERE / f).read_bytes()).hexdigest() for f in FILES},
    }
    (HERE / "phien_ban.json").write_text(json.dumps(d, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"Đã ghi phien_ban.json cho bản {d['phien_ban']} (thẻ {d['the']}).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
