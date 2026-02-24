"""Page Object: Główny edytor dokumentów — nagłówek, menu, dialogi."""

from playwright.sync_api import Locator, Page, expect

from pages.base_page import BasePage


class DocumentEditorPage(BasePage):
    """Reprezentuje główną stronę edytora dokumentów D2."""

    # ── Selektory ──────────────────────────────

    # Kontener
    CONTAINER = ".document-editor-container"

    # Header / Menu
    HEADER = "header.editor-header"
    APP_NAME = ".app-name"
    MENU_ITEMS = ".menu-item"

    # Menu rozwijane
    DROPDOWN_MENU = ".dropdown-menu"
    DROPDOWN_ITEM = ".dropdown-item"

    # Edytor
    EDITOR_MAIN = ".editor-main"
    PAPER_CONTAINER = ".paper-container"
    EDITOR_CONTENT = ".editor-content[contenteditable='true']"

    # Pasek statusu
    FOOTER = "footer.editor-footer"
    ZOOM_SLIDER = ".zoom-slider"
    PAGE_INDICATOR = ".page-indicator"

    # Komunikaty
    LOADING_OVERLAY = ".loading-overlay"
    ERROR_MESSAGE = ".error-message"
    SUCCESS_MESSAGE = ".success-message"

    # Dialogi
    ABOUT_DIALOG = ".about-dialog"
    TEMPLATES_DIALOG = ".templates-dialog"
    TEMPLATE_CARD = ".template-card"
    PAGE_SETUP_DIALOG = ".page-setup-dialog"
    HEADER_FOOTER_DIALOG = ".header-footer-dialog"
    FIND_REPLACE_DIALOG = ".find-replace-dialog"
    PARAGRAPH_DIALOG = ".paragraph-dialog"
    INSERT_TABLE_DIALOG = ".insert-table-dialog"
    PROPERTIES_DIALOG = ".properties-dialog"
    SIGNATURE_DIALOG = ".signature-dialog"

    # Kontekstowe
    CONTEXT_MENU = ".context-menu"
    CONTEXT_MENU_ITEM = ".context-menu-item"

    def __init__(self, page: Page) -> None:
        super().__init__(page)

    # ── Nawigacja ──────────────────────────────

    def open(self, base_url: str) -> "DocumentEditorPage":
        self.navigate(base_url)
        self.wait_for_visible(self.CONTAINER)
        return self

    def is_loaded(self) -> bool:
        return self.is_visible(self.CONTAINER)

    # ── Menu główne ────────────────────────────

    def open_menu(self, menu_name: str) -> None:
        """Otwiera menu po nazwie (np. 'Plik', 'Edytuj', 'Wstaw')."""
        self.page.locator(self.MENU_ITEMS, has_text=menu_name).click()

    def click_menu_item(self, menu_name: str, item_text: str) -> None:
        """Otwiera menu i klika pozycję."""
        self.open_menu(menu_name)
        self.page.locator(self.DROPDOWN_ITEM, has_text=item_text).click()

    # ── Operacje na dokumencie ─────────────────

    def new_document(self) -> None:
        self.click_menu_item("Plik", "Nowy dokument")

    def open_document_via_menu(self) -> None:
        self.click_menu_item("Plik", "Otwórz")

    def save_document(self) -> None:
        self.click_menu_item("Plik", "Zapisz")

    def save_document_as(self) -> None:
        self.click_menu_item("Plik", "Zapisz jako")

    # ── Edycja ─────────────────────────────────

    def undo(self) -> None:
        self.click_menu_item("Edytuj", "Cofnij")

    def redo(self) -> None:
        self.click_menu_item("Edytuj", "Ponów")

    # ── Edytor WYSIWYG ────────────────────────

    def get_editor(self) -> Locator:
        return self.page.locator(self.EDITOR_CONTENT).first

    def type_text(self, text: str) -> None:
        editor = self.get_editor()
        editor.click()
        editor.press_sequentially(text)

    def get_editor_html(self) -> str:
        return self.get_editor().inner_html()

    def get_editor_text(self) -> str:
        return self.get_editor().inner_text()

    def clear_editor(self) -> None:
        editor = self.get_editor()
        editor.click()
        self.page.keyboard.press("Control+A")
        self.page.keyboard.press("Delete")

    # ── Zoom ───────────────────────────────────

    def get_zoom_level(self) -> str:
        return self.get_text(self.PAGE_INDICATOR)

    # ── Dialogi ────────────────────────────────

    def open_page_setup(self) -> None:
        self.click_menu_item("Plik", "Ustawienia strony")
        self.wait_for_visible(self.PAGE_SETUP_DIALOG)

    def open_find_replace(self) -> None:
        self.click_menu_item("Edytuj", "Znajdź i zamień")

    def open_about(self) -> None:
        self.click_menu_item("Narzędzia", "O programie")
        self.wait_for_visible(self.ABOUT_DIALOG)

    def open_templates(self) -> None:
        self.click_menu_item("Plik", "Szablony")
        self.wait_for_visible(self.TEMPLATES_DIALOG)

    def select_template(self, template_name: str) -> None:
        self.page.locator(self.TEMPLATE_CARD, has_text=template_name).click()

    def open_properties(self) -> None:
        self.click_menu_item("Plik", "Właściwości")
        self.wait_for_visible(self.PROPERTIES_DIALOG)

    # ── Context menu ───────────────────────────

    def open_context_menu(self) -> None:
        self.get_editor().click(button="right")
        self.wait_for_visible(self.CONTEXT_MENU)

    def click_context_item(self, text: str) -> None:
        self.page.locator(self.CONTEXT_MENU_ITEM, has_text=text).click()

    # ── Asercje ────────────────────────────────

    def should_be_loaded(self) -> None:
        self.should_be_visible(self.CONTAINER)
        self.should_be_visible(self.HEADER)
        self.should_be_visible(self.EDITOR_MAIN)

    def should_show_loading(self) -> None:
        self.should_be_visible(self.LOADING_OVERLAY)

    def should_show_error(self, message: str | None = None) -> None:
        self.should_be_visible(self.ERROR_MESSAGE)
        if message:
            self.should_contain_text(self.ERROR_MESSAGE, message)

    def should_show_success(self, message: str | None = None) -> None:
        self.should_be_visible(self.SUCCESS_MESSAGE)
        if message:
            self.should_contain_text(self.SUCCESS_MESSAGE, message)

    def editor_should_contain(self, text: str) -> None:
        expect(self.get_editor()).to_contain_text(text)

    def editor_should_be_empty(self) -> None:
        content = self.get_editor_text().strip()
        assert content == "", f"Edytor nie jest pusty: '{content}'"
