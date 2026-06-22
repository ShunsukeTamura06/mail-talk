"""PyInstaller exe のエントリポイント（M端末で実行）。

`pyinstaller mailtalk.spec` でこのファイルを起点に単一exeを生成する。exe内に
Pythonランタイム・pywin32・FastAPI/uvicorn・static/index.html を内包するため、
M端末にPythonが無くても動く（CLAUDE.md §14）。
"""

from __future__ import annotations

import multiprocessing

from mailtalk.main import main

if __name__ == "__main__":
    # PyInstaller配下での多重起動防止（uvicorn/将来のワーカー対策）。
    multiprocessing.freeze_support()
    main()
