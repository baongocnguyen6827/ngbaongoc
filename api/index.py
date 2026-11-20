# -*- coding: utf-8 -*-
import os
import sys
from pathlib import Path

# Thêm đường dẫn thư mục gốc và thư mục làm việc hiện tại vào sys.path
root_dir = str(Path(__file__).resolve().parent.parent)
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)
if os.getcwd() not in sys.path:
    sys.path.insert(0, os.getcwd())

from app.main import app

if __name__ == "__main__":
    app.run()
