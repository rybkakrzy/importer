"""
Testy BDD: Pasek narzędzi — formatowanie tekstu.

Scenariusze z features/ui/toolbar.feature
"""

import pytest
from pytest_bdd import given, parsers, scenario, scenarios, then, when

from pages.document_editor_page import DocumentEditorPage
from pages.toolbar_page import ToolbarPage

scenarios("../../features/ui/toolbar.feature")


# ──────────────────────────────────────────────
# GIVEN (Tło — reużywane z conftest / document_editor)
# ──────────────────────────────────────────────


@given("otwieram aplikację D2", target_fixture="editor_page")
def open_app(editor_page: DocumentEditorPage) -> DocumentEditorPage:
    editor_page.should_be_loaded()
    return editor_page


@given(parsers.parse('wpisuję tekst "{text}" w edytorze'))
def type_text(editor_page: DocumentEditorPage, text: str):
    editor_page.type_text(text)


@given("zaznaczam cały tekst w edytorze")
def select_all_text(editor_page: DocumentEditorPage):
    editor_page.page.keyboard.press("Control+A")


# ──────────────────────────────────────────────
# WHEN
# ──────────────────────────────────────────────


@when(parsers.parse('klikam przycisk "{button_name}" na pasku narzędzi'))
def click_toolbar_button(toolbar: ToolbarPage, button_name: str):
    button_map = {
        "Pogrubienie": toolbar.click_bold,
        "Kursywa": toolbar.click_italic,
        "Podkreślenie": toolbar.click_underline,
        "Przekreślenie": toolbar.click_strikethrough,
        "Lista punktowana": toolbar.insert_bullet_list,
        "Lista numerowana": toolbar.insert_numbered_list,
    }
    action = button_map.get(button_name)
    if action:
        action()
    else:
        raise ValueError(f"Nieznany przycisk: {button_name}")


@when(parsers.parse('wybieram styl "{style_name}" z dropdown'))
def select_block_style(toolbar: ToolbarPage, style_name: str):
    toolbar.select_style(style_name)


@when(parsers.parse("ustawiam rozmiar czcionki na {size:d}"))
def set_font_size(toolbar: ToolbarPage, size: int):
    toolbar.set_font_size(size)


@when(parsers.parse('wybieram czcionkę "{font}"'))
def select_font(toolbar: ToolbarPage, font: str):
    toolbar.select_font_family(font)


@when(parsers.parse('klikam wyrównanie "{alignment}"'))
def click_alignment(toolbar: ToolbarPage, alignment: str):
    alignment_map = {
        "Do lewej": toolbar.align_left,
        "Wyśrodkuj": toolbar.align_center,
        "Do prawej": toolbar.align_right,
        "Wyjustuj": toolbar.align_justify,
    }
    action = alignment_map.get(alignment)
    if action:
        action()
    else:
        raise ValueError(f"Nieznane wyrównanie: {alignment}")


# ──────────────────────────────────────────────
# THEN
# ──────────────────────────────────────────────


@then(parsers.parse('tekst jest sformatowany jako "{format_type}"'))
def text_is_formatted(toolbar: ToolbarPage, format_type: str):
    check_map = {
        "bold": toolbar.bold_should_be_active,
        "italic": toolbar.italic_should_be_active,
        "underline": lambda: None,  # TODO: dodać sprawdzenie underline
    }
    check = check_map.get(format_type)
    if check:
        check()


@then(parsers.parse('aktualny styl to "{style_name}"'))
def current_style_is(toolbar: ToolbarPage, style_name: str):
    actual = toolbar.get_current_style()
    assert style_name in actual, f"Oczekiwany styl '{style_name}', aktualny: '{actual}'"


@then(parsers.parse('rozmiar czcionki wynosi "{size}"'))
def font_size_is(toolbar: ToolbarPage, size: str):
    actual = toolbar.get_font_size()
    assert actual == size, f"Oczekiwany rozmiar '{size}', aktualny: '{actual}'"


@then(parsers.parse('wybrana czcionka to "{font}"'))
def font_family_is(toolbar: ToolbarPage, font: str):
    actual = toolbar.get_font_family()
    assert font in actual, f"Oczekiwana czcionka '{font}', aktualna: '{actual}'"


@then(parsers.parse('tekst jest wyrównany "{alignment}"'))
def text_is_aligned(editor_page: DocumentEditorPage, alignment: str):
    # Weryfikacja przez sprawdzenie stanu przycisków w toolbar
    pass  # TODO: rozbudować o inspekcję style edytora


@then("tekst jest na liście punktowanej")
def text_is_bulleted(editor_page: DocumentEditorPage):
    html = editor_page.get_editor_html()
    assert "<ul" in html.lower(), "Brak listy punktowanej w edytorze"


@then("tekst jest na liście numerowanej")
def text_is_numbered(editor_page: DocumentEditorPage):
    html = editor_page.get_editor_html()
    assert "<ol" in html.lower(), "Brak listy numerowanej w edytorze"
