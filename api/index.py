import os
import sys

# Thêm thư mục gốc vào sys.path để có thể import package 'app' trên môi trường Vercel
root_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from app.main import app

# Các biến entrypoint WSGI chuẩn mực cấp module (top-level) cho Vercel Python Runtime
handler = app
application = app
wsgi_app = app

if __name__ == "__main__":
    app.run()
