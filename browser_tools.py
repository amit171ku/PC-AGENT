"""
browser_tools.py - Comprehensive Playwright toolkit for AI agents.
Covers browsing, automation, scraping, testing, screenshots/PDFs, and safety.
"""

import json
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional
from urllib.parse import urljoin

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

# ---------- State & Lifecycle ----------
_state = {
    "pw": None,
    "browser": None,
    "context": None,
    "page": None,
    "console_errors": [],
    "failed_requests": [],
}

# Optional confirmation callback: fn(prompt_text: str) -> bool
_confirm_callback: Optional[Callable[[str], bool]] = None


def set_confirm_handler(fn: Callable[[str], bool]):
    """Register a confirmation hook (e.g., UI dialog or terminal prompt)."""
    global _confirm_callback
    _confirm_callback = fn


def _require_confirm(action: str) -> bool:
    if _confirm_callback:
        return _confirm_callback(action)
    return True


def get_page():
    """Lazily initialize and return the current Playwright Page."""
    pg = _state.get("page")
    br = _state.get("browser")

    try:
        if pg is not None and br is not None and br.is_connected() and not pg.is_closed():
            return pg
    except Exception:
        pass

    close_browser()

    pw = sync_playwright().start()
    # Note: Headless mode can be set to False for visible automation
    browser = pw.chromium.launch(headless=True)
    context = browser.new_context(
        viewport={"width": 1280, "height": 800},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    )
    page = context.new_page()

    # Track console errors & network failures for testing & debugging
    page.on("console", lambda msg: _state["console_errors"].append(msg.text) if msg.type == "error" else None)
    page.on("requestfailed", lambda req: _state["failed_requests"].append(f"{req.method} {req.url} - {req.failure}"))

    _state.update({"pw": pw, "browser": browser, "context": context, "page": page})
    return page


def close_browser():
    """Teardown browser session and reset tracking state."""
    try:
        if _state.get("browser"):
            _state["browser"].close()
        if _state.get("pw"):
            _state["pw"].stop()
    except Exception:
        pass
    _state.update({
        "pw": None,
        "browser": None,
        "context": None,
        "page": None,
        "console_errors": [],
        "failed_requests": [],
    })


# =====================================================================
# 1. Web Browsing
# =====================================================================

def browser_open(url: str) -> str:
    """Navigate to a URL and wait for DOM content to load."""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    try:
        p = get_page()
        resp = p.goto(url, wait_until="domcontentloaded", timeout=30000)
        status = resp.status if resp else "unknown"
        return f"Loaded: {p.title()} (Status: {status}) - URL: {p.url}"
    except PlaywrightError as e:
        return f"Navigation Error: {e}"


def browser_navigate(action: str) -> str:
    """Navigate page history: 'back', 'forward', or 'reload'."""
    p = get_page()
    action = action.lower().strip()
    try:
        if action == "back":
            p.go_back(wait_until="domcontentloaded", timeout=15000)
        elif action == "forward":
            p.go_forward(wait_until="domcontentloaded", timeout=15000)
        elif action == "reload":
            p.reload(wait_until="domcontentloaded", timeout=15000)
        else:
            return f"Invalid navigation action: '{action}'. Use 'back', 'forward', or 'reload'."
        return f"Navigated {action}. Current page: {p.title()} ({p.url})"
    except PlaywrightError as e:
        return f"Navigation Error: {e}"


def browser_wait(timeout_ms: int = 2000, selector: Optional[str] = None) -> str:
    """Wait for a CSS selector to appear or pause for a specified duration."""
    p = get_page()
    try:
        if selector:
            p.wait_for_selector(selector, timeout=timeout_ms)
            return f"Found element matching '{selector}'."
        p.wait_for_timeout(timeout_ms)
        return f"Waited for {timeout_ms}ms."
    except PlaywrightError as e:
        return f"Wait failed: {e}"


# =====================================================================
# 2. Automation (Actions & Forms)
# =====================================================================

def browser_click(selector: str) -> str:
    """Click an element identified by CSS or XPath selector."""
    p = get_page()
    if not _require_confirm(f"Click '{selector}' on {p.url}"):
        return "Action cancelled by user."
    try:
        p.click(selector, timeout=10000)
        return f"Clicked '{selector}'."
    except PlaywrightError as e:
        return f"Click Error: {e}"


def browser_fill(selector: str, text: str, clear_first: bool = True) -> str:
    """Fill text into an input field or textarea."""
    # Safety guard: Prevent typing into sensitive fields
    lowered = selector.lower()
    if any(s in lowered for s in ("password", "passwd", "cvv", "card", "ssn", "secret")):
        return "Blocked by safety filter: refusing to input into sensitive fields."

    p = get_page()
    if not _require_confirm(f"Fill '{text}' into '{selector}'"):
        return "Action cancelled by user."

    try:
        if clear_first:
            p.fill(selector, text, timeout=10000)
        else:
            p.type(selector, text, timeout=10000)
        return f"Filled input '{selector}'."
    except PlaywrightError as e:
        return f"Fill Error: {e}"


def browser_select(selector: str, value: str) -> str:
    """Select an option in a <select> element by label or value."""
    p = get_page()
    if not _require_confirm(f"Select '{value}' on '{selector}'"):
        return "Action cancelled by user."
    try:
        p.select_option(selector, value=value, timeout=10000)
        return f"Selected '{value}' in '{selector}'."
    except PlaywrightError as e:
        return f"Select Error: {e}"


def browser_check(selector: str, checked: bool = True) -> str:
    """Check or uncheck a checkbox or radio button."""
    p = get_page()
    if not _require_confirm(f"{'Check' if checked else 'Uncheck'} '{selector}'"):
        return "Action cancelled by user."
    try:
        if checked:
            p.check(selector, timeout=10000)
        else:
            p.uncheck(selector, timeout=10000)
        return f"Element '{selector}' set to checked={checked}."
    except PlaywrightError as e:
        return f"Check Error: {e}"


def browser_press_key(key: str) -> str:
    """Press a keyboard key (e.g., 'Enter', 'Escape', 'Tab', 'ArrowDown')."""
    p = get_page()
    try:
        p.keyboard.press(key)
        return f"Pressed key '{key}'."
    except PlaywrightError as e:
        return f"Key Press Error: {e}"


def browser_scroll(direction: str = "down", amount: int = 600) -> str:
    """Scroll the active page up or down."""
    p = get_page()
    delta_y = amount if direction.lower() == "down" else -amount
    try:
        p.mouse.wheel(0, delta_y)
        p.wait_for_timeout(500)
        return f"Scrolled {direction} by {amount}px."
    except PlaywrightError as e:
        return f"Scroll Error: {e}"


def browser_hover(selector: str) -> str:
    """Hover over an element to trigger dropdowns or tooltips."""
    p = get_page()
    try:
        p.hover(selector, timeout=10000)
        return f"Hovered over '{selector}'."
    except PlaywrightError as e:
        return f"Hover Error: {e}"


# =====================================================================
# 3. Web Scraping
# =====================================================================

def browser_extract_text(selector: str = "body", max_length: int = 3500) -> str:
    """Extract visible inner text from an element or the entire page body."""
    p = get_page()
    try:
        text = p.inner_text(selector, timeout=10000).strip()
        if len(text) > max_length:
            return text[:max_length] + f"\n... [Truncated, total {len(text)} chars]"
        return text or "(empty)"
    except PlaywrightError as e:
        return f"Extraction Error: {e}"


def browser_extract_links(base_filter: Optional[str] = None, limit: int = 50) -> str:
    """Extract hyperlinks from the page, returning text and URLs."""
    p = get_page()
    try:
        links = p.eval_on_selector_all(
            "a[href]",
            """elements => elements.map(el => ({
                text: el.innerText.trim(),
                href: el.href
            })).filter(x => x.href && !x.href.startsWith('javascript:'))"""
        )
        if base_filter:
            links = [lk for lk in links if base_filter in lk["href"]]
        links = links[:limit]
        formatted = [f"- [{lk['text'] or 'link'}]({lk['href']})" for lk in links]
        return "\n".join(formatted) or "No matching links found."
    except PlaywrightError as e:
        return f"Extraction Error: {e}"


def browser_extract_table(selector: str = "table", max_rows: int = 25) -> str:
    """Extract an HTML <table> into structured Markdown table format."""
    p = get_page()
    try:
        table_data = p.eval_on_selector(
            selector,
            """table => {
                const rows = Array.from(table.querySelectorAll('tr'));
                return rows.map(r => Array.from(r.querySelectorAll('th, td')).map(c => c.innerText.trim()));
            }"""
        )
        if not table_data:
            return "Table is empty or not found."

        # Format as Markdown
        headers = table_data[0]
        divider = ["---"] * len(headers)
        lines = ["| " + " | ".join(headers) + " |", "| " + " | ".join(divider) + " |"]

        for row in table_data[1:max_rows]:
            # Pad row if columns don't align
            padded = row + [""] * (len(headers) - len(row))
            lines.append("| " + " | ".join(padded[:len(headers)]) + " |")

        if len(table_data) > max_rows:
            lines.append(f"... ({len(table_data) - max_rows} additional rows omitted)")

        return "\n".join(lines)
    except PlaywrightError as e:
        return f"Table Extraction Error: {e}"


# =====================================================================
# 4. Testing & Assertions
# =====================================================================

def browser_assert_element(selector: str, state: str = "visible", timeout_ms: int = 5000) -> str:
    """Assert an element exists in a given state ('visible', 'hidden', 'attached')."""
    p = get_page()
    try:
        p.wait_for_selector(selector, state=state, timeout=timeout_ms)
        return f"PASS: Element '{selector}' is {state}."
    except PlaywrightError:
        return f"FAIL: Element '{selector}' failed assertion (not {state} within {timeout_ms}ms)."


def browser_assert_text(expected_text: str, selector: str = "body") -> str:
    """Assert that the specified text is present inside an element or page."""
    p = get_page()
    try:
        content = p.inner_text(selector, timeout=5000)
        if expected_text in content:
            return f"PASS: Found text '{expected_text}' inside '{selector}'."
        return f"FAIL: Text '{expected_text}' not found in '{selector}'."
    except PlaywrightError as e:
        return f"Assertion Error: {e}"


def browser_get_console_errors() -> str:
    """Return collected browser console JavaScript errors and failed network requests."""
    errors = _state.get("console_errors", [])
    failed = _state.get("failed_requests", [])
    report = []
    if errors:
        report.append("Console Errors:\n" + "\n".join(f"- {e}" for e in errors[-10:]))
    if failed:
        report.append("Failed Network Requests:\n" + "\n".join(f"- {f}" for f in failed[-10:]))
    return "\n\n".join(report) if report else "No console errors or failed requests recorded."


# =====================================================================
# 5. Screenshots & PDFs
# =====================================================================

def browser_screenshot(path: str = "reports/screenshot.png", full_page: bool = True, selector: Optional[str] = None) -> str:
    """Capture a screenshot of the entire page or a target element."""
    p = get_page()
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        if selector:
            el = p.wait_for_selector(selector, timeout=10000)
            if not el:
                return f"Selector '{selector}' not found for screenshot."
            el.screenshot(path=str(target))
        else:
            p.screenshot(path=str(target), full_page=full_page)
        return f"Screenshot saved to: {target.as_posix()}"
    except PlaywrightError as e:
        return f"Screenshot Error: {e}"


def browser_export_pdf(path: str = "reports/page.pdf", format_size: str = "A4") -> str:
    """Export the active page as a PDF file."""
    p = get_page()
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        p.pdf(path=str(target), format=format_size, print_background=True)
        return f"PDF exported to: {target.as_posix()}"
    except PlaywrightError as e:
        return f"PDF Export Error (ensure browser launched in headless mode): {e}"


# =====================================================================
# 6. LLM Tool Registry & Schema Definitions
# =====================================================================

def _make_schema(name: str, desc: str, props: Dict[str, Any], required: List[str]) -> Dict[str, Any]:
    return {
        "type": "function",
        "function": {
            "name": name,
            "description": desc,
            "parameters": {
                "type": "object",
                "properties": props,
                "required": required,
            },
        },
    }


BROWSER_SCHEMAS = [
    _make_schema("browser_open", "Open a website by URL", {"url": {"type": "string", "description": "Web address to open"}}, ["url"]),
    _make_schema("browser_navigate", "Navigate page history: 'back', 'forward', or 'reload'", {"action": {"type": "string", "enum": ["back", "forward", "reload"]}}, ["action"]),
    _make_schema("browser_wait", "Wait for an element or pause", {"timeout_ms": {"type": "integer", "description": "Time in ms"}, "selector": {"type": "string", "description": "CSS selector to wait for"}}, []),
    _make_schema("browser_click", "Click a webpage element", {"selector": {"type": "string", "description": "CSS selector to click"}}, ["selector"]),
    _make_schema("browser_fill", "Type text into an input or form", {"selector": {"type": "string", "description": "CSS selector of field"}, "text": {"type": "string", "description": "Text to enter"}}, ["selector", "text"]),
    _make_schema("browser_select", "Select an option from a dropdown", {"selector": {"type": "string", "description": "Select CSS selector"}, "value": {"type": "string", "description": "Option value"}}, ["selector", "value"]),
    _make_schema("browser_check", "Toggle a checkbox or radio button", {"selector": {"type": "string", "description": "CSS selector"}, "checked": {"type": "boolean", "description": "True to check, False to uncheck"}}, ["selector"]),
    _make_schema("browser_press_key", "Press a key on the keyboard", {"key": {"type": "string", "description": "Key name like Enter, Tab, Escape"}}, ["key"]),
    _make_schema("browser_scroll", "Scroll up or down on the page", {"direction": {"type": "string", "enum": ["down", "up"]}, "amount": {"type": "integer", "description": "Pixel scroll amount"}}, []),
    _make_schema("browser_hover", "Hover mouse over an element", {"selector": {"type": "string", "description": "CSS selector"}}, ["selector"]),
    _make_schema("browser_extract_text", "Extract readable text from page or element", {"selector": {"type": "string", "description": "CSS selector (default 'body')"}, "max_length": {"type": "integer", "description": "Maximum characters"}}, []),
    _make_schema("browser_extract_links", "Extract hyperlinks from the page", {"base_filter": {"type": "string", "description": "Substring to filter URLs"}, "limit": {"type": "integer", "description": "Max links to return"}}, []),
    _make_schema("browser_extract_table", "Extract HTML table to structured markdown", {"selector": {"type": "string", "description": "CSS selector for table"}}, []),
    _make_schema("browser_assert_element", "Test if element is visible or hidden", {"selector": {"type": "string", "description": "CSS selector"}, "state": {"type": "string", "enum": ["visible", "hidden"]}}, ["selector"]),
    _make_schema("browser_assert_text", "Test if expected text is present", {"expected_text": {"type": "string", "description": "Text to find"}, "selector": {"type": "string", "description": "CSS selector"}}, ["expected_text"]),
    _make_schema("browser_get_console_errors", "Check for console JavaScript errors and failed requests", {}, []),
    _make_schema("browser_screenshot", "Capture page or element screenshot", {"path": {"type": "string", "description": "File path (e.g. reports/shot.png)"}, "full_page": {"type": "boolean", "description": "Whether to capture full page"}, "selector": {"type": "string", "description": "Optional element selector"}}, []),
    _make_schema("browser_export_pdf", "Export page as PDF file", {"path": {"type": "string", "description": "Target PDF path"}}, []),
]

BROWSER_FUNCS = {
    "browser_open": browser_open,
    "browser_navigate": browser_navigate,
    "browser_wait": browser_wait,
    "browser_click": browser_click,
    "browser_fill": browser_fill,
    "browser_select": browser_select,
    "browser_check": browser_check,
    "browser_press_key": browser_press_key,
    "browser_scroll": browser_scroll,
    "browser_hover": browser_hover,
    "browser_extract_text": browser_extract_text,
    "browser_extract_links": browser_extract_links,
    "browser_extract_table": browser_extract_table,
    "browser_assert_element": browser_assert_element,
    "browser_assert_text": browser_assert_text,
    "browser_get_console_errors": browser_get_console_errors,
    "browser_screenshot": browser_screenshot,
    "browser_export_pdf": browser_export_pdf,
}