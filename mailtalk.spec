# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec（B端末でビルド → A端末で実行する単一exe）。

ビルド: B端末(Windows,Python)で
    uv pip install -e ".[windows,build]"
    pyinstaller mailtalk.spec
生成物: dist/MailTalk.exe（共有フォルダ経由でA端末へ持ち込み、ダブルクリック）

win32com・uvicorn・pydanticは動的importが多くPyInstallerが取りこぼすため、
collect_submodulesで明示的に集める。static/index.htmlはdatasで同梱する。
"""

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = (
    collect_submodules("uvicorn")
    + collect_submodules("fastapi")
    + collect_submodules("pydantic")
    + collect_submodules("anyio")
    + collect_submodules("win32com")
    + [
        "win32timezone",
        "pythoncom",
        "pywintypes",
    ]
)

a = Analysis(
    ["run_app.py"],
    pathex=["src"],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy"],
    noarchive=False,
)

# static配下を丸ごと同梱（将来CSS/JS/画像を分割しても欠落しない）。
a.datas += Tree("src/mailtalk/static", prefix="static")

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="MailTalk",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
