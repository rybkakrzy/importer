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
