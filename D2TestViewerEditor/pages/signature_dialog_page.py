"""Page Object: Dialog podpisywania dokumentu certyfikatem cyfrowym."""

from playwright.sync_api import Page, expect

from pages.base_page import BasePage


class SignatureDialogPage(BasePage):
    """Dialog podpisywania dokumentu (POST /document/sign)."""

    # ── Selektory ──────────────────────────────

    DIALOG = ".signature-dialog"
    DIALOG_OVERLAY = ".dialog-overlay"

    # Pola formularza
    SIGNER_NAME_INPUT = "input#signerName"
    SIGNER_TITLE_INPUT = "input#signerTitle"
    SIGNER_EMAIL_INPUT = "input#signerEmail"
    REASON_INPUT = "input#signatureReason, textarea#signatureReason"
    CERTIFICATE_INPUT = "input#certificateFile"
    PASSWORD_INPUT = "input#certificatePassword"

    # Przyciski
    SIGN_BTN = ".signature-dialog .btn-primary"
    CANCEL_BTN = ".signature-dialog .btn-secondary"

    # Status / walidacja
    VALIDATION_ERROR = ".signature-dialog .validation-error, .signature-dialog .field-error"
    SUCCESS_TOAST = ".toast.success, [role='alert'].success"
    LOADING_SPINNER = ".signature-dialog .spinner, .signature-dialog .loading"

    def __init__(self, page: Page) -> None:
        super().__init__(page)

    # ── Interakcje ─────────────────────────────

    def wait_for_open(self) -> None:
        self.wait_for_visible(self.DIALOG)

    def fill_signer_name(self, name: str) -> None:
        self.fill(self.SIGNER_NAME_INPUT, name)

    def fill_signer_title(self, title: str) -> None:
        self.fill(self.SIGNER_TITLE_INPUT, title)

    def fill_signer_email(self, email: str) -> None:
        self.fill(self.SIGNER_EMAIL_INPUT, email)

    def fill_reason(self, reason: str) -> None:
        self.page.locator(self.REASON_INPUT).first.fill(reason)

    def fill_certificate_password(self, password: str) -> None:
        self.fill(self.PASSWORD_INPUT, password)

    def upload_certificate(self, file_path: str) -> None:
        self.page.locator(self.CERTIFICATE_INPUT).set_input_files(file_path)

    def click_sign(self) -> None:
        self.click(self.SIGN_BTN)

    def click_cancel(self) -> None:
        self.click(self.CANCEL_BTN)

    # ── Asercje ────────────────────────────────

    def should_be_open(self) -> None:
        super().should_be_visible(self.DIALOG)

    def should_be_visible(self) -> None:
        """Zgodność z konwencją BasePage."""
        self.should_be_open()

    def should_be_closed(self) -> None:
        self.wait_for_hidden(self.DIALOG)

    def sign_button_should_be_enabled(self) -> None:
        expect(self.page.locator(self.SIGN_BTN)).to_be_enabled()

    def sign_button_should_be_disabled(self) -> None:
        expect(self.page.locator(self.SIGN_BTN)).to_be_disabled()

    def should_show_validation_error(self, message: str | None = None) -> None:
        super().should_be_visible(self.VALIDATION_ERROR)
        if message:
            self.should_contain_text(self.VALIDATION_ERROR, message)
