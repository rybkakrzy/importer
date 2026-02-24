"""Page Object: Dialog generowania kodów kreskowych / QR."""

from playwright.sync_api import Page, expect

from pages.base_page import BasePage


class BarcodeDialogPage(BasePage):
    """Dialog wstawiania kodu kreskowego i kodu QR."""

    # ── Selektory ──────────────────────────────

    DIALOG_OVERLAY = ".dialog-overlay"
    DIALOG = ".dialog.barcode-dialog"
    TYPE_SELECT = "select#barcodeType"
    CONTENT_INPUT = "input#barcodeContent"
    WIDTH_INPUT = "input#barcodeWidth"
    HEIGHT_INPUT = "input#barcodeHeight"
    SHOW_VALUE_CHECKBOX = ".form-group-checkbox input[type='checkbox']"
    PREVIEW_IMAGE = "img.barcode-preview-img"
    PREVIEW_ERROR = ".preview-error"
    PREVIEW_PLACEHOLDER = ".preview-placeholder"
    INSERT_BTN = ".dialog .btn-primary"
    CANCEL_BTN = ".dialog .btn-secondary"

    def __init__(self, page: Page) -> None:
        super().__init__(page)

    # ── Interakcje ─────────────────────────────

    def wait_for_open(self) -> None:
        self.wait_for_visible(self.DIALOG)

    def select_type(self, barcode_type: str) -> None:
        self.select_option(self.TYPE_SELECT, barcode_type)

    def enter_content(self, content: str) -> None:
        self.fill(self.CONTENT_INPUT, content)

    def set_width(self, width: int) -> None:
        self.fill(self.WIDTH_INPUT, str(width))

    def set_height(self, height: int) -> None:
        self.fill(self.HEIGHT_INPUT, str(height))

    def toggle_show_value(self) -> None:
        self.click(self.SHOW_VALUE_CHECKBOX)

    def click_insert(self) -> None:
        self.click(self.INSERT_BTN)

    def click_cancel(self) -> None:
        self.click(self.CANCEL_BTN)

    def close_via_overlay(self) -> None:
        self.click(self.DIALOG_OVERLAY)

    # ── Asercje ────────────────────────────────

    def should_be_open(self) -> None:
        super().should_be_visible(self.DIALOG)

    def should_be_closed(self) -> None:
        self.wait_for_hidden(self.DIALOG)

    def should_show_preview(self) -> None:
        super().should_be_visible(self.PREVIEW_IMAGE)

    def should_show_placeholder(self) -> None:
        super().should_be_visible(self.PREVIEW_PLACEHOLDER)

    def should_show_error(self, message: str | None = None) -> None:
        super().should_be_visible(self.PREVIEW_ERROR)
        if message:
            self.should_contain_text(self.PREVIEW_ERROR, message)

    def insert_should_be_disabled(self) -> None:
        expect(self.page.locator(self.INSERT_BTN)).to_be_disabled()

    def insert_should_be_enabled(self) -> None:
        expect(self.page.locator(self.INSERT_BTN)).to_be_enabled()
