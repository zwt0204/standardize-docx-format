#!/usr/bin/env python3
"""Lightweight regression checks for TOC-safe paragraph matching."""
from __future__ import annotations

import importlib.util
from pathlib import Path


def load_apply_profile():
    path = Path(__file__).with_name("apply_profile.py")
    spec = importlib.util.spec_from_file_location("apply_profile", path)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def main() -> None:
    mod = load_apply_profile()
    body = {
        "style": "2",
        "styleName": "heading 1",
        "paragraph": None,
        "inTable": False,
        "inTextBox": False,
        "section": 2,
        "index": 10,
    }
    toc = {
        "style": "7",
        "styleName": "toc 1",
        "paragraph": None,
        "inTable": False,
        "inTextBox": False,
        "section": 1,
        "index": 3,
    }

    # Monkeypatch paragraph_text via info["paragraph"] path by patching function
    def fake_paragraph_text(paragraph):
        return paragraph

    mod.paragraph_text = fake_paragraph_text  # type: ignore
    body["paragraph"] = "1 绪论"
    toc["paragraph"] = "1 绪论 1"

    body_match = {
        "textRegex": r"^\d+\s+\S+",
        "styleNameNotRegex": r"(?i)^toc",
        "textNotRegex": r".+\d$",
    }
    assert mod.paragraph_matches(body, body_match) is True
    assert mod.paragraph_matches(toc, body_match) is False
    assert mod.paragraph_matches(toc, {"styleNameRegex": r"(?i)^toc"}) is True
    assert mod.paragraph_matches(body, {"currentStyleNotIn": ["7", "8"]}) is True
    assert mod.paragraph_matches(toc, {"currentStyleNotIn": ["7", "8"]}) is False
    print("paragraph match safety: 5 checks passed")


if __name__ == "__main__":
    main()
