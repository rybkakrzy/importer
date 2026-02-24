"""
Testy BDD: Dialog kodu kreskowego.

Scenariusze z features/ui/barcode.feature
"""

import re

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from pages.barcode_dialog_page import BarcodeDialogPage
from pages.document_editor_page import DocumentEditorPage

scenarios("../../features/ui/barcode.feature")


# ──────────────────────────────────────────────
# GIVEN (Tło)
# ──────────────────────────────────────────────


@given("otwieram aplikację D2", target_fixture="editor_page")
def open_app(editor_page: DocumentEditorPage) -> DocumentEditorPage:
    editor_page.should_be_loaded()
    return editor_page


@given("otwieram dialog kodu kreskowego")
def open_barcode_dialog(editor_page: DocumentEditorPage, barcode_dialog: BarcodeDialogPage):
    editor_page.open_menu("Wstaw")
    editor_page.page.get_by_text("Kod kreskowy").click()
    barcode_dialog.should_be_visible()


# ──────────────────────────────────────────────
# WHEN
# ──────────────────────────────────────────────


@when(parsers.parse('wybieram typ kodu "{barcode_type}"'))
def select_barcode_type(barcode_dialog: BarcodeDialogPage, barcode_type: str):
    barcode_dialog.select_type(barcode_type)


@when(parsers.parse('wpisuję zawartość "{content}"'))
def type_barcode_content(barcode_dialog: BarcodeDialogPage, content: str):
    barcode_dialog.set_content(content)


@when(parsers.parse('klikam "{button_label}"'))
def click_dialog_button(barcode_dialog: BarcodeDialogPage, button_label: str):
    if button_label == "Wstaw":
        barcode_dialog.click_insert()
    elif button_label == "Anuluj":
        barcode_dialog.click_cancel()
    else:
        raise ValueError(f"Nieznany przycisk: {button_label}")


@when(parsers.parse("ustawiam szerokość na {width:d}"))
def set_width(barcode_dialog: BarcodeDialogPage, width: int):
    barcode_dialog.set_width(width)


@when(parsers.parse("ustawiam wysokość na {height:d}"))
def set_height(barcode_dialog: BarcodeDialogPage, height: int):
    barcode_dialog.set_height(height)


# ──────────────────────────────────────────────
# THEN
# ──────────────────────────────────────────────


@then("dialog kodu kreskowego jest widoczny")
def dialog_visible(barcode_dialog: BarcodeDialogPage):
    barcode_dialog.should_be_visible()


@then("dialog kodu kreskowego jest zamknięty")
def dialog_closed(barcode_dialog: BarcodeDialogPage):
    barcode_dialog.should_be_closed()


@then(parsers.parse('przycisk "{button}" jest nieaktywny'))
def button_disabled(barcode_dialog: BarcodeDialogPage, button: str):
    barcode_dialog.insert_button_should_be_disabled()


@then(parsers.parse('przycisk "{button}" jest aktywny'))
def button_enabled(barcode_dialog: BarcodeDialogPage, button: str):
    barcode_dialog.insert_button_should_be_enabled()


@then("podgląd kodu kreskowego jest widoczny")
def preview_visible(barcode_dialog: BarcodeDialogPage):
    barcode_dialog.preview_should_be_visible()


@then("edytor zawiera obraz kodu kreskowego")
def editor_contains_barcode_image(editor_page: DocumentEditorPage):
    html = editor_page.get_editor_html()
    assert re.search(r"<img[^>]+barcode", html, re.IGNORECASE) or \
           "data:image" in html, \
        "Edytor nie zawiera obrazu kodu kreskowego"
