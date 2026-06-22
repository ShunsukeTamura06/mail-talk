"""テスト共通設定。設定キャッシュを毎テスト既定値に固定し、決定的にする。"""

from __future__ import annotations

import pytest

from mailtalk import config


@pytest.fixture(autouse=True)
def _default_config():
    """各テストの前後で Config を既定値に固定（config.json の影響を排除）。"""
    config.set_config(config.Config())
    yield
    config.reset_config()
