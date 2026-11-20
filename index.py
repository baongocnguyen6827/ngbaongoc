import os
import sys

# Thêm thư mục gốc vào sys.path
root_dir = os.path.dirname(os.path.abspath(__file__))
if root_dir not in sys.path:
    sys.path.insert(0, root_dir)

from app.main import app

# Top-level WSGI entrypoints cho Vercel Python Runtime
handler = app
application = app
wsgi_app = app

if __name__ == "__main__":
    app.run()

