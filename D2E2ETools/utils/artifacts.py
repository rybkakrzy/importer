"""
Helpery do screenshotów i artefaktów testowych.
"""

from __future__ import annotations

import datetime
from pathlib import Path

from playwright.sync_api import Page

ARTIFACTS_DIR = Path(__file__).resolve().parent.parent / "test-results" / "artifacts"


def take_screenshot(page: Page, name: str | None = None) -> Path:
    """Wykonuje screenshot i zwraca ścieżkę do pliku."""
    ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"{name or 'screenshot'}_{timestamp}.png"
    path = ARTIFACTS_DIR / filename
    page.screenshot(path=str(path), full_page=True)
    return path


def save_page_html(page: Page, name: str | None = None) -> Path:
    """Zapisuje HTML strony jako artefakt debugowania."""
    ARTIFACTS_DIR.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"{name or 'page'}_{timestamp}.html"
    path = ARTIFACTS_DIR / filename
    path.write_text(page.content(), encoding="utf-8")
    return path
