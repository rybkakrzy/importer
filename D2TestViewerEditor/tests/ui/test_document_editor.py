"""
Testy BDD: Edytor dokumentów — podstawowe operacje.

Scenariusze z features/ui/document_editor.feature
"""

import pytest
from pytest_bdd import given, parsers, scenario, scenarios, then, when

from pages.document_editor_page import DocumentEditorPage
from pages.toolbar_page import ToolbarPage

# Załaduj wszystkie scenariusze z feature file
scenarios("../../features/ui/document_editor.feature")


# ──────────────────────────────────────────────
# GIVEN (Tło)
# ──────────────────────────────────────────────


@given("otwieram aplikację D2")
def open_app(editor_page: DocumentEditorPage):
    """Aplikacja jest już otwarta dzięki fixture app_page → editor_page."""
    editor_page.should_be_loaded()


# ──────────────────────────────────────────────
# WHEN
# ──────────────────────────────────────────────


@when(parsers.parse('wpisuję tekst "{text}" w edytorze'))
def type_text_in_editor(editor_page: DocumentEditorPage, text: str):
    editor_page.type_text(text)


@when(parsers.parse('wybieram z menu "{menu}" opcję "{option}"'))
def select_menu_option(editor_page: DocumentEditorPage, menu: str, option: str):
    editor_page.click_menu_item(menu, option)


@when("otwieram dialog szablonów")
def open_templates_dialog(editor_page: DocumentEditorPage):
    editor_page.open_templates()


@when(parsers.parse('wybieram szablon "{template_name}"'))
def select_template(editor_page: DocumentEditorPage, template_name: str):
    editor_page.select_template(template_name)


@when("wykonuję cofnij")
def perform_undo(editor_page: DocumentEditorPage):
    editor_page.page.keyboard.press("Control+Z")


@when("wykonuję ponów")
def perform_redo(editor_page: DocumentEditorPage):
    editor_page.page.keyboard.press("Control+Y")


@when("klikam prawym przyciskiem w obszarze edytora")
def right_click_editor(editor_page: DocumentEditorPage):
    editor_page.open_context_menu()


# ──────────────────────────────────────────────
# THEN
# ──────────────────────────────────────────────


@then(parsers.parse('widzę nagłówek aplikacji z napisem "{text}"'))
def see_app_header(editor_page: DocumentEditorPage, text: str):
    editor_page.should_contain_text(DocumentEditorPage.APP_NAME, text)


@then("widzę pasek narzędzi edytora")
def see_toolbar(toolbar: ToolbarPage):
    toolbar.should_be_visible()


@then("widzę obszar edycji dokumentu")
def see_editor_area(editor_page: DocumentEditorPage):
    editor_page.should_be_visible(DocumentEditorPage.EDITOR_CONTENT)


@then(parsers.parse('edytor zawiera tekst "{text}"'))
def editor_contains_text(editor_page: DocumentEditorPage, text: str):
    editor_page.editor_should_contain(text)


@then("edytor jest pusty")
def editor_is_empty(editor_page: DocumentEditorPage):
    editor_page.editor_should_be_empty()


@then("inicjowane jest pobieranie pliku DOCX")
def download_initiated(editor_page: DocumentEditorPage):
    """Sprawdza czy Playwright zarejestrował zdarzenie pobierania pliku."""
    # Oczekujemy, że po kliknięciu Zapisz serwer zwróci plik do pobrania
    with editor_page.page.expect_download(timeout=10_000) as download_info:
        editor_page.save_document()
    download = download_info.value
    assert download.suggested_filename.endswith(".docx"), (
        f"Pobrano plik o nieoczekiwanej nazwie: {download.suggested_filename}"
    )


@then("widoczna jest stopka edytora")
def see_editor_footer(editor_page: DocumentEditorPage):
    editor_page.should_be_visible(DocumentEditorPage.FOOTER)


@then("widoczny jest wskaźnik strony")
def see_page_indicator(editor_page: DocumentEditorPage):
    editor_page.should_be_visible(DocumentEditorPage.PAGE_INDICATOR)


@then("menu kontekstowe jest widoczne")
def context_menu_visible(editor_page: DocumentEditorPage):
    editor_page.should_be_visible(DocumentEditorPage.CONTEXT_MENU)
