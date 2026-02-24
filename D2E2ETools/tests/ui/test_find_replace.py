"""
Testy BDD: Wyszukiwanie i zamiana tekstu.

Scenariusze z features/ui/find_replace.feature
"""

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from pages.document_editor_page import DocumentEditorPage

scenarios("../../features/ui/find_replace.feature")


# ──────────────────────────────────────────────
# GIVEN (Tło)
# ──────────────────────────────────────────────


@given("otwieram aplikację D2", target_fixture="editor_page")
def open_app(editor_page: DocumentEditorPage) -> DocumentEditorPage:
    editor_page.should_be_loaded()
    return editor_page


@given(parsers.parse('wpisuję tekst "{text}" w edytorze'))
def type_text(editor_page: DocumentEditorPage, text: str):
    editor_page.type_text(text)


# ──────────────────────────────────────────────
# WHEN
# ──────────────────────────────────────────────


@when("otwieram pasek wyszukiwania")
def open_search_bar(editor_page: DocumentEditorPage):
    """Otwiera panel wyszukiwania skrótem Ctrl+H lub przez menu."""
    editor_page.open_menu("Edytuj")
    editor_page.page.get_by_text("Znajdź i zamień").click()


@when(parsers.parse('wyszukuję tekst "{search_text}"'))
def search_for_text(editor_page: DocumentEditorPage, search_text: str):
    search_input = editor_page.page.locator('[data-testid="find-input"], input[placeholder*="Szukaj"], .find-replace input').first
    search_input.fill(search_text)
    search_input.press("Enter")


# ──────────────────────────────────────────────
# THEN
# ──────────────────────────────────────────────


@then("pasek wyszukiwania jest widoczny")
def search_bar_visible(editor_page: DocumentEditorPage):
    search_panel = editor_page.page.locator(
        '[data-testid="find-replace-panel"], .find-replace, .find-bar'
    ).first
    search_panel.wait_for(state="visible", timeout=5_000)


@then(parsers.parse("liczba wyników wyszukiwania jest większa niż {count:d}"))
def results_greater_than(editor_page: DocumentEditorPage, count: int):
    result_label = editor_page.page.locator(
        '[data-testid="search-results-count"], .search-count, .find-replace .match-count'
    ).first
    result_label.wait_for(state="visible", timeout=5_000)
    text = result_label.inner_text()
    # Oczekiwany format: "1 z 3", "2/5", "3 wyniki", itp.
    import re
    digits = re.findall(r"\d+", text)
    total = int(digits[-1]) if digits else 0
    assert total > count, f"Oczekiwano >  {count} wyników, znaleziono: {total} (tekst: '{text}')"


@then(parsers.parse("liczba wyników wyszukiwania wynosi {count:d}"))
def results_equal(editor_page: DocumentEditorPage, count: int):
    result_label = editor_page.page.locator(
        '[data-testid="search-results-count"], .search-count, .find-replace .match-count'
    ).first
    result_label.wait_for(state="visible", timeout=5_000)
    text = result_label.inner_text()
    import re
    digits = re.findall(r"\d+", text)
    total = int(digits[-1]) if digits else 0
    assert total == count, f"Oczekiwano {count} wyników, znaleziono: {total} (tekst: '{text}')"
