"""
Testy BDD: Podpisywanie dokumentów cyfrowym certyfikatem.

Scenariusze z features/ui/signature.feature
"""

from __future__ import annotations

import pytest
from pytest_bdd import given, parsers, scenarios, then, when

from pages.document_editor_page import DocumentEditorPage
from pages.signature_dialog_page import SignatureDialogPage

scenarios("../../features/ui/signature.feature")


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


@when("otwieram dialog podpisywania")
def open_sign_dialog(editor_page: DocumentEditorPage, signature_dialog: SignatureDialogPage):
    editor_page.click_menu_item("Narzędzia", "Podpisz dokument")
    signature_dialog.wait_for_open()


@when(parsers.parse('klikam "Anuluj" w dialogu podpisywania'))
def click_cancel_sign(signature_dialog: SignatureDialogPage):
    signature_dialog.click_cancel()


@when(parsers.parse('wpisuję imię podpisującego "{name}"'))
def fill_signer_name(signature_dialog: SignatureDialogPage, name: str):
    signature_dialog.fill_signer_name(name)


@when(parsers.parse('wpisuję tytuł podpisującego "{title}"'))
def fill_signer_title(signature_dialog: SignatureDialogPage, title: str):
    signature_dialog.fill_signer_title(title)


@when(parsers.parse('wpisuję email podpisującego "{email}"'))
def fill_signer_email(signature_dialog: SignatureDialogPage, email: str):
    signature_dialog.fill_signer_email(email)


@when(parsers.parse('wpisuję powód podpisania "{reason}"'))
def fill_sign_reason(signature_dialog: SignatureDialogPage, reason: str):
    signature_dialog.fill_reason(reason)


# ──────────────────────────────────────────────
# THEN
# ──────────────────────────────────────────────


@then("dialog podpisywania jest widoczny")
def sign_dialog_visible(signature_dialog: SignatureDialogPage):
    signature_dialog.should_be_visible()


@then("dialog podpisywania jest zamknięty")
def sign_dialog_closed(signature_dialog: SignatureDialogPage):
    signature_dialog.should_be_closed()


@then(parsers.parse('przycisk "{button}" jest nieaktywny'))
def sign_button_disabled(signature_dialog: SignatureDialogPage, button: str):
    signature_dialog.sign_button_should_be_disabled()


@then("przycisk \"Podpisz\" jest nadal nieaktywny bez certyfikatu")
def sign_button_still_disabled(signature_dialog: SignatureDialogPage):
    signature_dialog.sign_button_should_be_disabled()


@then("pole email wskazuje błąd walidacji")
def email_validation_error(signature_dialog: SignatureDialogPage):
    # Weryfikacja przez sprawdzenie atrybutu HTML5 lub klasy błędu
    email_input = signature_dialog.page.locator(SignatureDialogPage.SIGNER_EMAIL_INPUT)
    validity = email_input.evaluate("el => el.validity.valid")
    assert not validity, "Pole email nie zgłasza błędu walidacji dla nieprawidłowego adresu"
