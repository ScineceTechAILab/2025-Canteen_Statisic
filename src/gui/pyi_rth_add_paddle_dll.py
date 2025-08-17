
import os, sys
from pathlib import Path

def _add_dll_dir(p):
    if hasattr(os, 'add_dll_directory'):
        os.add_dll_directory(str(p))
    else:
        os.environ['PATH'] = str(p) + os.pathsep + os.environ.get('PATH', '')

# 运行在打包环境（_MEIPASS）时，DLL 都在这个目录附近
base = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))

# 常见几种布局都尝试一下
candidates = [
    base / 'paddle.libs',
    base / '_internal' / 'paddle.libs',
    base / 'paddle' / 'libs',
]

for c in candidates:
    if c.exists():
        _add_dll_dir(c)
