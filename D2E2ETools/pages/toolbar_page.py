"""Page Object: Pasek narzędzi edytora (toolbar)."""

from playwright.sync_api import Locator, Page, expect

from pages.base_page import BasePage


class ToolbarPage(BasePage):
    """Pasek narzędzi — formatowanie, czcionki, wstawianie elementów."""

    # ── Selektory ──────────────────────────────

    TOOLBAR = ".editor-toolbar"
    TOOLBAR_BTN = ".toolbar-btn"

    # Blok / styl
    STYLE_DROPDOWN = ".style-dropdown"
    STYLE_DROPDOWN_TRIGGER = ".style-dropdown-trigger"
    STYLE_DROPDOWN_MENU = ".style-dropdown-menu"
    STYLE_OPTION = ".style-option"

    # Czcionka
    FONT_FAMILY_SELECT = "select.toolbar-select.font-family-select"
    FONT_SIZE_INPUT = "input.font-size-input"
    FONT_SIZE_INCREASE = ".font-size-btn:first-child"
    FONT_SIZE_DECREASE = ".font-size-btn:last-child"

    # Kolory
    TEXT_COLOR_INPUT = ".color-picker-wrapper input.color-input[type='color']"

    # Wyszukiwanie
    SEARCH_BAR = ".search-bar"
    SEARCH_INPUT = "input.search-input"
    SEARCH_COUNT = ".search-count"
    REPLACE_ROW = ".replace-row"

    # Link dialog
    LINK_DIALOG = ".dialog-overlay"
    LINK_URL_INPUT = "#linkUrl"
    LINK_TEXT_INPUT = "#linkText"

    def __init__(self, page: Page) -> None:
        super().__init__(page)

    # ── Formatowanie tekstu ────────────────────

    def click_bold(self) -> None:
        self.page.locator(self.TOOLBAR_BTN, has_text="B").first.click()

    def click_italic(self) -> None:
        self.page.locator(self.TOOLBAR_BTN, has_text="I").first.click()

    def click_underline(self) -> None:
        self.page.locator(self.TOOLBAR_BTN, has_text="U").first.click()

    def click_strikethrough(self) -> None:
        self.page.get_by_title("Przekreślenie").click()

    def is_bold_active(self) -> bool:
        btn = self.page.locator(self.TOOLBAR_BTN, has_text="B").first
        return "active" in (btn.get_attribute("class") or "")

    def is_italic_active(self) -> bool:
        btn = self.page.locator(self.TOOLBAR_BTN, has_text="I").first
        return "active" in (btn.get_attribute("class") or "")

    # ── Styl bloku ─────────────────────────────

    def open_style_dropdown(self) -> None:
        self.click(self.STYLE_DROPDOWN_TRIGGER)
        self.wait_for_visible(self.STYLE_DROPDOWN_MENU)

    def select_style(self, style_name: str) -> None:
        self.open_style_dropdown()
        self.page.locator(self.STYLE_OPTION, has_text=style_name).click()

    def get_current_style(self) -> str:
        return self.get_text(self.STYLE_DROPDOWN_TRIGGER)

    # ── Czcionka ───────────────────────────────

    def select_font_family(self, font: str) -> None:
        self.select_option(self.FONT_FAMILY_SELECT, font)

    def get_font_family(self) -> str:
        return self.get_input_value(self.FONT_FAMILY_SELECT)

    def set_font_size(self, size: int) -> None:
        inp = self.page.locator(self.FONT_SIZE_INPUT)
        inp.fill(str(size))
        inp.press("Enter")

    def get_font_size(self) -> str:
        return self.get_input_value(self.FONT_SIZE_INPUT)

    def increase_font_size(self) -> None:
        self.page.get_by_title("Zwiększ czcionkę").click()

    def decrease_font_size(self) -> None:
        self.page.get_by_title("Zmniejsz czcionkę").click()

    # ── Wyrównanie ─────────────────────────────

    def align_left(self) -> None:
        self.page.get_by_title("Do lewej").click()

    def align_center(self) -> None:
        self.page.get_by_title("Wyśrodkuj").click()

    def align_right(self) -> None:
        self.page.get_by_title("Do prawej").click()

    def align_justify(self) -> None:
        self.page.get_by_title("Wyjustuj").click()

    # ── Listy ──────────────────────────────────

    def insert_bullet_list(self) -> None:
        self.page.get_by_title("Lista punktowana").click()

    def insert_numbered_list(self) -> None:
        self.page.get_by_title("Lista numerowana").click()

    # ── Wstawianie ─────────────────────────────

    def open_insert_link(self) -> None:
        self.page.get_by_title("Wstaw link").click()
        self.wait_for_visible(self.LINK_DIALOG)

    def insert_link(self, url: str, text: str = "") -> None:
        self.open_insert_link()
        self.fill(self.LINK_URL_INPUT, url)
        if text:
            self.fill(self.LINK_TEXT_INPUT, text)
        self.page.locator(".dialog .btn-primary").click()

    def click_insert_image(self) -> None:
        self.page.get_by_title("Wstaw obraz").click()

    def click_insert_table(self) -> None:
        self.page.get_by_title("Wstaw tabelę").click()

    def click_insert_barcode(self) -> None:
        self.page.get_by_title("Kod kreskowy").click()

    # ── Undo / Redo ────────────────────────────

    def click_undo(self) -> None:
        self.page.get_by_title("Cofnij").click()

    def click_redo(self) -> None:
        self.page.get_by_title("Ponów").click()

    # ── Szukaj / Zamień ────────────────────────

    def toggle_search(self) -> None:
        self.page.get_by_title("Znajdź").click()

    def search_text(self, text: str) -> None:
        self.toggle_search()
        self.wait_for_visible(self.SEARCH_BAR)
        self.fill(self.SEARCH_INPUT, text)

    def get_search_results_count(self) -> str:
        return self.get_text(self.SEARCH_COUNT)

    # ── Asercje ────────────────────────────────

    def should_be_visible(self, selector: str | None = None) -> None:
        super().should_be_visible(selector or self.TOOLBAR)

    def bold_should_be_active(self) -> None:
        assert self.is_bold_active(), "Przycisk Bold nie jest aktywny"

    def italic_should_be_active(self) -> None:
        assert self.is_italic_active(), "Przycisk Italic nie jest aktywny"
