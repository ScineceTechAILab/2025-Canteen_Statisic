# -*- mode: python ; coding: utf-8 -*-

import os, sys, glob, site, pathlib
from PyInstaller.utils.hooks import (
    collect_dynamic_libs,
    collect_submodules,
    collect_data_files,
)
import paddleocr


# ===== hidden imports（尽量保守，后续可精简）=====
hidden_imports = list(set(
    [
        'paddle',
        'cv2',
        'numpy',
        'pandas',
        'openpyxl',
        'xlrd',
        'xlwt',
        'xlwings',
        'shutil',
        'multiprocessing',
        'threading',
        'subprocess',
        'datetime',
        'configparser',
        'PySide6',
        'PySide6.QtCore',
        'PySide6.QtGui',
        'PySide6.QtWidgets',
        'sklearn',
        'matplotlib',
        'matplotlib.pyplot',
        'Levenshtein',
        'albumentations',
        'Cython',  # 导入到 Cython
        # Paddle/PaddleOCR 常见入口
        'paddle.base.core',
        'paddle.base.libpaddle',
        'paddle.fluid.core_avx',
        'paddle.fluid.libpaddle',
    ] 
))

# ===== 强制收集 paddle 的 DLL =====
def find_paddle_libs():
    candidates = []
    try:
        candidates.extend(site.getsitepackages())
    except Exception:
        pass
    try:
        usp = site.getusersitepackages()
        if usp: candidates.append(usp)
    except Exception:
        pass

    dll_pairs = []
    for root in candidates:
        p = os.path.join(root, 'paddle', 'libs')
        if os.path.isdir(p):
            for dll in glob.glob(os.path.join(p, '*.dll')):
                dll_pairs.append((dll, 'paddle.libs'))
            break
    return dll_pairs

paddle_lib_dlls = find_paddle_libs()
auto_paddle_bins = collect_dynamic_libs("paddle")

# OpenCV 常见 ffmpeg 动态库
opencv_bins = []
try:
    import cv2
    cv2_dir = os.path.dirname(cv2.__file__)
    opencv_bins = [(p, '.') for p in glob.glob(os.path.join(cv2_dir, 'opencv_videoio_ffmpeg*.dll'))]
except Exception:
    pass

binaries = []
binaries += paddle_lib_dlls
binaries += auto_paddle_bins
binaries += opencv_bins

# ===== 数据文件：PaddleOCR缓存 + Cython 资源（关键！）=====

ppocr = os.path.join(os.path.dirname(paddleocr.__file__), 'ppocr')

datas = [
    ('../.paddleocr', 'src/.paddleocr'), # 将 main_window.py 所在目录的上级目录下的 .paddleocr 目录加入输出目录的 src 目录下
    ('../data', 'src/data'),             # 将 main_window.py 所在目录的上级目录下的 data 目录加入输出目录的 src 目录下
    (ppocr,'ppocr')                      # 把 PaddleOCR 的 ppocr 数据加入输出目录的 ppocr 目录下
]

# 把 Cython 的资源（尤其是 Utility 模板）全部打进包
datas += collect_data_files('Cython')  # 若想更小，也可只收 Utility：collect_data_files('Cython', includes=['Utility/*'])

# 把 tools 目录加入数据
import paddle
import paddleocr
import shutil
from PyInstaller.building.datastruct import TOC

tools_src = os.path.join(os.path.dirname(paddleocr.__file__), "tools")
tools_dst = "paddleocr/tools"
if os.path.exists(tools_src):
    datas += [(tools_src, "tools")]

# ===== 运行时 Hook：把 paddle.libs 加入 DLL 搜索路径 =====
_runtime_hook_code = r"""
import os, sys
from pathlib import Path

def _add_dll_dir(p):
    if hasattr(os, 'add_dll_directory'):
        os.add_dll_directory(str(p))
    else:
        os.environ['PATH'] = str(p) + os.pathsep + os.environ.get('PATH', '')

base = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent))

candidates = [
    base / 'paddle.libs',
    base / '_internal' / 'paddle.libs',
    base / 'paddle' / 'libs',
]

for c in candidates:
    if c.exists():
        _add_dll_dir(c)
"""

_rth_path = os.path.abspath('pyi_rth_add_paddle_dll.py')
with open(_rth_path, 'w', encoding='utf-8') as f:
    f.write(_runtime_hook_code)

a = Analysis(
    ['main_window.py'],
    pathex=['..'],
    binaries=binaries,
    datas=datas,
    
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[_rth_path],
    excludes=[
        'paddle.jit.sot',
    ],
    noarchive=False,  # 如仍遇到资源查找问题，可尝试 True（不打 zip）
    optimize=0,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=False,
    name='main_window',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,     # 避免压缩破坏 DLL
    console=True,  # 仅 GUI 可改 False
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    contents_directory='.',   # Mistake：禁用 _internal 布局使得 exe 与资源文件打包在同一目录下
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='main_window',
)
