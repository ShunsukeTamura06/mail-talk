"""診断バンドル収集のテスト（macOS/Fakeで完結検証）。"""

from __future__ import annotations

import json
import zipfile

from mailtalk.diagnostics import collect_diagnostics
from mailtalk.fake_outlook import FakeOutlookSource
from mailtalk.models import LANE_ORDER


def _read_report(zip_path) -> dict:
    with zipfile.ZipFile(zip_path) as z:
        return json.loads(z.read("diagnostics.json"))


def test_collect_produces_zip_with_report(tmp_path):
    zip_path = collect_diagnostics(FakeOutlookSource(), out_dir=tmp_path)
    assert zip_path.exists()
    report = _read_report(zip_path)
    assert report["summary"]["conversations"] >= 4
    assert set(report["summary"]["lane_counts"]) == set(LANE_ORDER)
    assert report["self_identification"]["count"] == 1
    assert report["errors"] == []


def test_no_mail_body_in_bundle(tmp_path):
    # 本文(body)は持ち帰りに含めない。
    zip_path = collect_diagnostics(FakeOutlookSource(), out_dir=tmp_path)
    with zipfile.ZipFile(zip_path) as z:
        raw = z.read("diagnostics.json").decode("utf-8")
    assert "body_preview" not in raw
    assert "body_html" not in raw


def test_redact_masks_addresses_and_subjects(tmp_path):
    zip_path = collect_diagnostics(FakeOutlookSource(), out_dir=tmp_path, redact=True)
    report = _read_report(zip_path)
    assert report["redacted"] is True
    # 自分アドレス me@bank.co.jp が生で出ない（マスクされる）。
    addrs = report["self_identification"]["my_addresses"]
    assert all("***@" in x for x in addrs)
    # 件名は "(件名 len=...)" 形式。
    for c in report["conversations"]:
        assert c["subject_norm"].startswith("(件名 len=")
