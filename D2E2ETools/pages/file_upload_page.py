"""Page Object: Komponent przesyłania plików ZIP."""

from pathlib import Path

from playwright.sync_api import Page, expect

from pages.base_page import BasePage


class FileUploadPage(BasePage):
    """Strona/komponent przesyłania archiwów ZIP."""

    # ── Selektory ──────────────────────────────

    CONTAINER = ".upload-container"
    FILE_INPUT = "input#fileInput"
    FILE_LABEL = "label.file-label"
    FILE_NAME = ".file-name"
    UPLOAD_BTN = "button.upload-button"
    ERROR_MESSAGE = ".error-message"
    RESPONSE_SECTION = ".response-section"
    SUCCESS_HEADER = ".success-header"
    SUMMARY_GRID = ".summary-grid"
    FILES_LIST = ".files-list"
    FILE_ITEM = ".file-item"
    EMPTY_ARCHIVE = ".empty-archive"

    def __init__(self, page: Page) -> None:
        super().__init__(page)

    # ── Interakcje ─────────────────────────────

    def upload_file(self, file_path: str | Path) -> None:
        self.page.locator(self.FILE_INPUT).set_input_files(str(file_path))

    def click_upload(self) -> None:
        self.click(self.UPLOAD_BTN)

    def get_file_name(self) -> str:
        return self.get_text(self.FILE_NAME)

    def get_files_count(self) -> int:
        return self.page.locator(self.FILE_ITEM).count()

    # ── Asercje ────────────────────────────────

    def should_show_success(self) -> None:
        self.wait_for_visible(self.SUCCESS_HEADER)

    def should_show_error(self, message: str | None = None) -> None:
        super().should_be_visible(self.ERROR_MESSAGE)
        if message:
            self.should_contain_text(self.ERROR_MESSAGE, message)

    def should_show_files(self, count: int | None = None) -> None:
        super().should_be_visible(self.FILES_LIST)
        if count is not None:
            expect(self.page.locator(self.FILE_ITEM)).to_have_count(count)

    def should_show_empty_archive(self) -> None:
        super().should_be_visible(self.EMPTY_ARCHIVE)
