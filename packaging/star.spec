# -*- mode: python ; coding: utf-8 -*-
# 星易查 桌面版 PyInstaller spec（onedir 模式，Windows / Linux / macOS）
#
# 构建（在仓库根目录执行）：
#   pyinstaller packaging/star.spec --noconfirm --distpath dist --workpath build/pyinstaller
#
# 产物（按平台）：
#   Windows: dist/星易查/星易查.exe   -> Inno Setup 打成 Setup.exe
#   Linux:   dist/XingYiCha/XingYiCha -> AppImage / tar.gz 组装
#   macOS:   dist/XingYiCha/XingYiCha.app（Contents/Frameworks 为 onedir 内容）
#
# 选 onedir 而非 onefile：OCR 栈（onnxruntime + opencv + pymupdf + rapidocr 模型）
# 打包后约 500MB，onefile 每次启动需先解压到临时目录，启动慢且更易被杀软误报。

import os
import sys

from PyInstaller.utils.hooks import collect_all

PROJECT = os.path.abspath(os.path.join(SPECPATH, '..'))
ON_MACOS = sys.platform == 'darwin'
ON_WINDOWS = os.name == 'nt'

# 产物名：Windows 用中文名（已验证可用）；Linux/macOS 用 ASCII 名，
# 避免 AppImage / dmg / 上传工具链的非 ASCII 兼容问题。
NAME = '星易查' if ON_WINDOWS else 'XingYiCha'

datas = [
    (os.path.join(PROJECT, 'templates'), 'templates'),
    (os.path.join(PROJECT, 'static'), 'static'),
]
binaries = []
hiddenimports = ['waitress', 'cv2', 'fitz', 'onnxruntime']
if ON_MACOS:
    # app.py 冻结入口用 AppKit 跑 Cocoa 事件循环（Dock reopen / 菜单栏）
    hiddenimports += ['AppKit', 'Foundation']

# OCR 引擎两个发行包二选一：py<3.13 → rapidocr_onnxruntime；py>=3.13 →
# rapidocr（+onnxruntime）。两者都把 ONNX 模型与 YAML 配置作为 package
# data 分发，静态分析发现不了；app.py 又是在函数内动态 import 的，所以
# 对两个包各跑一次 collect_all 兜底收集。包缺失时 collect_all 返回空三元
# 组或抛异常都不影响构建——运行时 OCR 自动降级跳过（app.py 的设计行为）。
for _ocr_pkg in ('rapidocr', 'rapidocr_onnxruntime'):
    try:
        _ocr_datas, _ocr_binaries, _ocr_hiddenimports = collect_all(_ocr_pkg)
        datas += _ocr_datas
        binaries += _ocr_binaries
        hiddenimports += _ocr_hiddenimports
    except Exception:
        pass

# pymupdf.layout（pymupdf-layout 发行包）同理：~50MB ONNX 模型 + yaml 配置
# 在 resources/onnx 下作为 package data 分发，且 app.py 在函数内动态 import，
# collect_all 兜底收集（包缺失时 collect_all 返回空三元组，不影响构建）。
try:
    _layout_datas, _layout_binaries, _layout_hiddenimports = collect_all('pymupdf.layout')
    datas += _layout_datas
    binaries += _layout_binaries
    hiddenimports += _layout_hiddenimports
except Exception:
    pass

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

if ON_MACOS:
    # macOS 模板顺序：EXE(bootloader) -> COLLECT(onedir) -> BUNDLE 包成 .app，
    # binaries/datas 进 Contents/Frameworks。CFBundleDisplayName 用中文，
    # Finder 中显示"星易查"。
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=True,  # Windows/Linux：控制台窗口看进度，关窗即停服；macOS Finder 启动无窗口，交互走 Cocoa 图形层（app.py _run_macos_gui）
        icon=os.path.join(SPECPATH, 'star.icns') if os.path.exists(os.path.join(SPECPATH, 'star.icns')) else None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        name=NAME,
    )
    app_bundle = BUNDLE(
        coll,
        name=f'{NAME}.app',
        icon=os.path.join(SPECPATH, 'star.icns') if os.path.exists(os.path.join(SPECPATH, 'star.icns')) else None,
        bundle_identifier='com.xingyicha.app',
        info_plist={
            'CFBundleDisplayName': '星易查',
            'CFBundleName': NAME,
            'CFBundleShortVersionString': '2.2.0',
            'NSHighResolutionCapable': True,
            # onnxruntime>=1.19 的 macOS x86_64 wheel 最低要求 10.15，
            # 标低了会在 10.13/10.14 上启动即 import 失败（OCR 降级）。
            'LSMinimumSystemVersion': '10.15',
        },
    )
else:
    exe = EXE(
        pyz,
        a.scripts,
        [],
        exclude_binaries=True,
        name=NAME,
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=False,
        console=True,  # 保留控制台窗口：看分析进度，关闭窗口即停止服务
        icon=os.path.join(SPECPATH, 'star.ico') if ON_WINDOWS and os.path.exists(os.path.join(SPECPATH, 'star.ico')) else None,
    )
    coll = COLLECT(
        exe,
        a.binaries,
        a.datas,
        strip=False,
        upx=False,
        name=NAME,
    )
