"""
Wspólne fixture'y i step definitions dla testów UI.

Importuje Page Objecty i udostępnia je jako fixture pytest.
"""

import pytest
from playwright.sync_api import Page

from config.settings import BASE_URL
from pages.barcode_dialog_page import BarcodeDialogPage
from pages.document_editor_page import DocumentEditorPage
from pages.toolbar_page import ToolbarPage


# ──────────────────────────────────────────────
# Page Object fixtures
# ──────────────────────────────────────────────


@pytest.fixture()
def editor_page(app_page: Page) -> DocumentEditorPage:
    """Zwraca Page Object edytora dokumentów (strona już załadowana)."""
    return DocumentEditorPage(app_page)


@pytest.fixture()
def toolbar(app_page: Page) -> ToolbarPage:
    """Zwraca Page Object paska narzędzi."""
    return ToolbarPage(app_page)


@pytest.fixture()
def barcode_dialog(app_page: Page) -> BarcodeDialogPage:
    """Zwraca Page Object dialogu kodów kreskowych."""
    return BarcodeDialogPage(app_page)
