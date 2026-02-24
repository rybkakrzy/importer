"""Bazowa klasa Page Object — wspólna logika dla wszystkich stron."""

from playwright.sync_api import Locator, Page, expect


class BasePage:
    """Wspólne metody dla wszystkich Page Objectów."""

    def __init__(self, page: Page) -> None:
        self.page = page

    # ── Nawigacja ──────────────────────────────

    def navigate(self, url: str) -> None:
        self.page.goto(url)

    def reload(self) -> None:
        self.page.reload()

    # ── Oczekiwanie ────────────────────────────

    def wait_for_visible(self, selector: str, timeout: int | None = None) -> Locator:
        locator = self.page.locator(selector)
        locator.wait_for(state="visible", timeout=timeout)
        return locator

    def wait_for_hidden(self, selector: str, timeout: int | None = None) -> None:
        self.page.locator(selector).wait_for(state="hidden", timeout=timeout)

    # ── Interakcje ─────────────────────────────

    def click(self, selector: str) -> None:
        self.page.locator(selector).click()

    def fill(self, selector: str, value: str) -> None:
        self.page.locator(selector).fill(value)

    def select_option(self, selector: str, value: str) -> None:
        self.page.locator(selector).select_option(value)

    def get_text(self, selector: str) -> str:
        return self.page.locator(selector).inner_text()

    def get_input_value(self, selector: str) -> str:
        return self.page.locator(selector).input_value()

    def is_visible(self, selector: str) -> bool:
        return self.page.locator(selector).is_visible()

    def is_enabled(self, selector: str) -> bool:
        return self.page.locator(selector).is_enabled()

    # ── Asercje (expect) ───────────────────────

    def should_be_visible(self, selector: str) -> None:
        expect(self.page.locator(selector)).to_be_visible()

    def should_have_text(self, selector: str, text: str) -> None:
        expect(self.page.locator(selector)).to_have_text(text)

    def should_contain_text(self, selector: str, text: str) -> None:
        expect(self.page.locator(selector)).to_contain_text(text)

    # ── Screenshot ─────────────────────────────

    def screenshot(self, path: str = "screenshot.png") -> bytes:
        return self.page.screenshot(path=path, full_page=True)
