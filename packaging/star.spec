# -*- mode: python ; coding: utf-8 -*-
# 星易查 Windows 桌面版 PyInstaller spec（onedir 模式）
#
# 构建（在仓库根目录执行）：
#   pyinstaller packaging/star.spec --noconfirm --distpath dist --workpath build/pyinstaller-win
#
# 产物：dist/星易查/星易查.exe（templates/static/OCR模型 均在其 _internal 下）
#
# 选 onedir 而非 onefile：OCR 栈（onnxruntime + opencv + pymupdf + rapidocr 模型）
# 打包后约 500MB，onefile 每次启动需先解压到临时目录，启动慢且更易被杀软误报。

import os

from PyInstaller.utils.hooks import collect_all

PROJECT = os.path.abspath(os.path.join(SPECPATH, '..'))
ICON = os.path.join(SPECPATH, 'star.ico')

datas = [
    (os.path.join(PROJECT, 'templates'), 'templates'),
    (os.path.join(PROJECT, 'static'), 'static'),
]
binaries = []
hiddenimports = ['waitress', 'cv2', 'fitz', 'onnxruntime']

# rapidocr_onnxruntime 把 ONNX 模型与 YAML 配置作为 package data 分发，
# 静态分析发现不了，collect_all 一次性收集 data 文件 / DLL / 子模块。
_ocr_datas, _ocr_binaries, _ocr_hiddenimports = collect_all('rapidocr_onnxruntime')
datas += _ocr_datas
binaries += _ocr_binaries
hiddenimports += _ocr_hiddenimports

a = Analysis(
    [os.path.join(PROJECT, 'app.py')],
    pathex=[PROJECT],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'IPython', 'pytest', 'gunicorn'],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='星易查',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,  # 保留控制台窗口：看分析进度，关闭窗口即停止服务
    icon=ICON if os.path.exists(ICON) else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    name='星易查',
)
