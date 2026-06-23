"""clean_body（本文正規化・非削除）のテスト。"""

from __future__ import annotations

from mailtalk.models import clean_body


def test_normalizes_crlf_to_lf():
    assert clean_body("a\r\nb\r\nc") == "a\nb\nc"
    assert clean_body("a\rb") == "a\nb"


def test_collapses_3plus_blank_lines():
    assert clean_body("a\n\n\n\nb") == "a\n\nb"


def test_strips_trailing_space_per_line():
    assert clean_body("a   \nb\t\n") == "a\nb"


def test_does_not_delete_quoted_or_reply_history():
    # この職場では `>` 付き行に回答本体を書く運用 → clean_body は消してはいけない。
    src = "Bさん\n回答しました\n===========\n質問\n> 1/1です\n> 日本橋です"
    out = clean_body(src)
    assert "1/1です" in out
    assert "日本橋です" in out
    assert "===========" in out


def test_empty_is_safe():
    assert clean_body("") == ""
    assert clean_body(None) == ""
