# -*- mode: python ; coding: utf-8 -*-

import os, sys, glob, site, pathlib

from PyInstaller.utils.hooks import (
    collect_dynamic_libs,
    collect_submodules,
    collect_data_files,
)

# ===== 1) hidden imports（可适当精简，但为了稳妥先保守一些） =====
# 补一手 OCR 的子模块
hidden_imports = list(set(
    [
        # 你原有的
        'paddleocr',
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
        'Cython',

        # Paddle/PaddleOCR 常见动态依赖
        'paddle.base.core',
        'paddle.base.libpaddle',
        'paddle.fluid.core_avx',
        'paddle.fluid.libpaddle',
    ]
    + collect_submodules('paddleocr')   # 这里放到 list 里面再 set 去重
))

# 如果包体积太大，成功后你再来精简

# ===== 2) 强制收集 Paddle 的 DLL（显式查找 paddle/libs/*.dll） =====
def find_paddle_libs():
    # 搜索所有可能的 site-packages 路径（兼容 venv / user-site）
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
    found_dir = None
    for root in candidates:
        p = os.path.join(root, 'paddle', 'libs')
        if os.path.isdir(p):
            found_dir = p
            # 把 paddle/libs 下所有 DLL 都拷到 dist 的 paddle.libs 目录中
            for dll in glob.glob(os.path.join(p, '*.dll')):
                dll_pairs.append((dll, 'paddle.libs'))
            break
    return found_dir, dll_pairs

paddle_lib_dir, paddle_lib_dlls = find_paddle_libs()

# 保险起见，仍然保留 PyInstaller 的自动收集结果（有些 .pyd 依赖能识别到）
auto_paddle_bins = collect_dynamic_libs("paddle")

# 你可能还会用到的：OpenCV 的 ffmpeg 动态库
opencv_bins = []
try:
    import cv2, os, glob
    cv2_dir = os.path.dirname(cv2.__file__)
    # 常见：opencv_videoio_ffmpeg*.dll
    opencv_bins = [(p, '.') for p in glob.glob(os.path.join(cv2_dir, 'opencv_videoio_ffmpeg*.dll'))]
except Exception:
    pass

# 合并所有二进制
binaries = []
# 显式收集到的 paddle DLL 放前面（避免被其他 hook 覆盖路径）
binaries += paddle_lib_dlls
binaries += auto_paddle_bins
binaries += opencv_bins

# ===== 3) 数据文件（你原来拷的 OCR 缓存/字典等） =====
datas = [
    ('../.paddleocr', '.paddleocr'),
]
# 如果需要也可以收集 paddle 的 data（通常不必须）
# datas += collect_data_files('paddle')

# ===== 4) 运行时 Hook：把 paddle.libs 加入 DLL 搜索路径 =====
# 解决某些机器上 Windows 对非系统目录 DLL 搜索受限的问题
_runtime_hook_code = r"""
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
"""

# 将 runtime hook 写入一个本地文件，供 Analysis 使用
_rth_path = os.path.abspath('pyi_rth_add_paddle_dll.py')
with open(_rth_path, 'w', encoding='utf-8') as f:
    f.write(_runtime_hook_code)

# ===== 5) Analysis / EXE / COLLECT =====
a = Analysis(
    ['main_window.py'],
    pathex=['..'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[_rth_path],   # ✅ 关键：确保运行时能找到 DLL
    excludes=[
        'paddle.jit.sot',        # 你原来的排除项
    ],
    noarchive=False,
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
    upx=False,                   # ✅ 关闭 UPX，避免压缩破坏 DLL
    console=True,                # 需要无控制台可改为 False
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,                   # 同样关闭
    upx_exclude=[],
    name='main_window',
)
