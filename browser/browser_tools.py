from typing import Any, Callable, Dict, List, Optional
from urllib.parse import quote_plus

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from common import ROOT, safe

# ---------- State & Lifecycle ----------
_state: Dict[str, Any] = {
    "pw": None,
    "browser": None,
    "context": None,
    "page": None,
    "console_errors": [],
    "failed_requests": [],
}

# Confirmation callback: fn(prompt_text) -> bool
_confirm_callback: Optional[Callable[[str], bool]] = None

SENSITIVE_WORDS = ("password", "passwd", "card", "cvv", "secret", "ssn")


def set_confirm_handler(fn: Callable[[str], bool]):
    """Register the confirmation hook (the web UI approval dialog)."""
    global _confirm_callback
    _confirm_callback = fn


def _require_confirm(action: str) -> bool:
    """Ask the user. Denies when no handler is registered (safe default)."""
    if _confirm_callback:
        return bool(_confirm_callback(action))
    return False


def _int(value, default: int) -> int:
    """LLMs often send numbers as strings; accept both."""
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return default


def _bool(value, default: bool = True) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return str(value).strip().lower() not in ("false", "0", "no", "off", "")


def get_page():
    """Lazily start Chromium and return the active Page (restarts if closed)."""
    pg = _state.get("page")
    br = _state.get("browser")
    try:
        if pg is not None and br is not None and br.is_connected() and not pg.is_closed():
            return pg
    except Exception:
        pass

    close_browser()

    pw = sync_playwright().start()
    browser = pw.chromium.launch(
        headless=False,  # visible window (as described in the README)
        args=[
            "--disable-blink-features=AutomationControlled",
            "--no-sandbox",
            "--disable-infobars",
            "--start-maximized",
        ],
    )
    context = browser.new_context(
        viewport={"width": 1280, "height": 800},
        user_agent=(
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        ),
    )
    page = context.new_page()
    page.add_init_script(
        "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
    )

    # Track console errors and failed requests for testing/debugging
    page.on(
        "console",
        lambda msg: _state["console_errors"].append(msg.text) if msg.type == "error" else None,
    )
    page.on(
        "requestfailed",
        lambda req: _state["failed_requests"].append(f"{req.method} {req.url} - {req.failure}"),
    )

    _state.update({"pw": pw, "browser": browser, "context": context, "page": page})
    return page


def close_browser():
    """Close the browser and reset all tracking state."""
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
# 1. Search & Browsing
# =====================================================================

def search_web(query: str) -> str:
    """Search Google (falls back to Bing if blocked) and return result text."""
    p = get_page()
    try:
        p.goto(
            f"https://www.google.com/search?q={quote_plus(query)}&hl=en",
            wait_until="domcontentloaded", timeout=25000,
        )

        # Dismiss cookie/consent banner if present
        for button_text in ("Accept all", "I agree", "Reject all", "Read all"):
            try:
                btn = p.locator(f"button:has-text('{button_text}')").first
                if btn.is_visible(timeout=1000):
                    btn.click()
                    p.wait_for_timeout(1000)
                    break
            except Exception:
                pass

        body_text = p.inner_text("body")
        blocked = any(
            phrase in body_text.lower()
            for phrase in ("unusual traffic", "recaptcha", "solve the challenge",
                           "bot detection", "robot")
        )

        if not blocked:
            try:
                p.wait_for_selector("#search", timeout=3000)
                results = p.locator("#search").inner_text()
                if results and len(results.strip()) > 100:
                    return results[:4000]
            except Exception:
                pass
            if len(body_text.strip()) > 200:
                return body_text[:4000]
    except Exception:
        pass

    # Silent fallback: Bing
    print(f"[search] Google challenged; falling back to Bing for: '{query}'")
    try:
        p.goto(
            f"https://www.bing.com/search?q={quote_plus(query)}&setlang=en",
            wait_until="domcontentloaded", timeout=25000,
        )
        try:
            results = p.locator("#b_results").inner_text(timeout=4000)
            if results and len(results.strip()) > 100:
                return results[:4000]
        except Exception:
            pass
        return p.inner_text("body")[:4000]
    except Exception as e:
        return f"Search error: {e}"


def browser_open(url: str) -> str:
    """Navigate to a URL and wait for the DOM to load."""
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    try:
        p = get_page()
        resp = p.goto(url, wait_until="domcontentloaded", timeout=30000)
        status = resp.status if resp else "unknown"
        return f"Opened: {p.title()} (Status: {status}) - {p.url}"
    except PlaywrightError as e:
        return f"Navigation Error: {e}"


def browser_navigate(action: str) -> str:
    """Navigate history: 'back', 'forward', or 'reload'."""
    p = get_page()
    act = action.lower().strip()
    try:
        if act == "back":
            p.go_back(wait_until="domcontentloaded", timeout=15000)
        elif act == "forward":
            p.go_forward(wait_until="domcontentloaded", timeout=15000)
        elif act == "reload":
            p.reload(wait_until="domcontentloaded", timeout=15000)
        else:
            return f"Invalid action: '{action}'. Use 'back', 'forward', or 'reload'."
        return f"Navigated {act}. Page: {p.title()} ({p.url})"
    except PlaywrightError as e:
        return f"Navigation Error: {e}"


def browser_wait(selector: str = "", timeout_ms: str = "2000") -> str:
    """Wait for a selector to appear, or pause for a duration."""
    p = get_page()
    ms = _int(timeout_ms, 2000)
    try:
        if selector and selector.strip():
            p.wait_for_selector(selector.strip(), timeout=ms)
            return f"Found '{selector}'."
        p.wait_for_timeout(ms)
        return f"Waited {ms}ms."
    except PlaywrightError as e:
        return f"Wait failed: {e}"


# =====================================================================
# 2. Automation (actions & forms)
# =====================================================================

def browser_click(selector: str) -> str:
    p = get_page()
    if not _require_confirm(f"Click: {selector} on {p.url}"):
        return "User denied."
    try:
        p.click(selector, timeout=10000)
        return f"Clicked '{selector}'."
    except PlaywrightError as e:
        return f"Click Error: {e}"


def browser_fill(selector: str, text: str, clear_first: str = "true") -> str:
    if any(w in selector.lower() for w in SENSITIVE_WORDS):
        return "Refused: safety filter prevents typing into sensitive fields."
    p = get_page()
    if not _require_confirm(f"Type: '{text}' into {selector}"):
        return "User denied."
    try:
        if _bool(clear_first):
            p.fill(selector, text, timeout=10000)
        else:
            p.locator(selector).press_sequentially(text, timeout=10000)
        return f"Typed into '{selector}'."
    except PlaywrightError as e:
        return f"Fill Error: {e}"


def browser_select(selector: str, value: str) -> str:
    p = get_page()
    if not _require_confirm(f"Select '{value}' in {selector}"):
        return "User denied."
    try:
        p.select_option(selector, value=value, timeout=10000)
        return f"Selected '{value}' in '{selector}'."
    except PlaywrightError as e:
        return f"Select Error: {e}"


def browser_check(selector: str, checked: str = "true") -> str:
    want = _bool(checked)
    p = get_page()
    if not _require_confirm(f"{'Check' if want else 'Uncheck'}: {selector}"):
        return "User denied."
    try:
        if want:
            p.check(selector, timeout=10000)
        else:
            p.uncheck(selector, timeout=10000)
        return f"Set '{selector}' checked={want}."
    except PlaywrightError as e:
        return f"Check Error: {e}"


def browser_press_key(key: str) -> str:
    p = get_page()
    try:
        p.keyboard.press(key)
        return f"Pressed '{key}'."
    except PlaywrightError as e:
        return f"Key Press Error: {e}"


def browser_scroll(direction: str = "down", amount: str = "600") -> str:
    p = get_page()
    px = _int(amount, 600)
    delta = px if direction.lower() == "down" else -px
    try:
        p.mouse.wheel(0, delta)
        p.wait_for_timeout(300)
        return f"Scrolled {direction} by {px}px."
    except PlaywrightError as e:
        return f"Scroll Error: {e}"


def browser_hover(selector: str) -> str:
    p = get_page()
    try:
        p.hover(selector, timeout=10000)
        return f"Hovered over '{selector}'."
    except PlaywrightError as e:
        return f"Hover Error: {e}"


# =====================================================================
# 3. Web scraping
# =====================================================================

def browser_extract_text(selector: str = "body", max_length: str = "4000") -> str:
    p = get_page()
    limit = _int(max_length, 4000)
    try:
        text = p.inner_text(selector or "body", timeout=10000).strip()
        if len(text) > limit:
            return text[:limit] + f"\n... [truncated, total {len(text)} chars]"
        return text or "(empty)"
    except PlaywrightError as e:
        return f"Extraction Error: {e}"


def browser_extract_links(filter_keyword: str = "", limit: str = "30") -> str:
    p = get_page()
    try:
        links = p.eval_on_selector_all(
            "a[href]",
            """els => els.map(a => ({
                text: a.innerText.trim().replace(/\\s+/g, ' '),
                href: a.href
            })).filter(x => x.href && !x.href.startsWith('javascript:'))""",
        )
        if filter_keyword and filter_keyword.strip():
            kw = filter_keyword.lower().strip()
            links = [l for l in links if kw in l["text"].lower() or kw in l["href"].lower()]
        links = links[:_int(limit, 30)]
        return "\n".join(f"- [{l['text'] or 'link'}]({l['href']})" for l in links) \
            or "No matching links found."
    except PlaywrightError as e:
        return f"Extraction Error: {e}"


def browser_extract_table(selector: str = "table", max_rows: str = "25") -> str:
    p = get_page()
    limit = _int(max_rows, 25)
    try:
        rows = p.eval_on_selector(
            selector or "table",
            """t => Array.from(t.querySelectorAll('tr')).map(r =>
                Array.from(r.querySelectorAll('th, td')).map(c => c.innerText.trim().replace(/\\s+/g, ' '))
            )""",
        )
        if not rows:
            return "Table is empty or not found."
        headers = rows[0]
        lines = [
            "| " + " | ".join(headers) + " |",
            "| " + " | ".join(["---"] * len(headers)) + " |",
        ]
        for r in rows[1:limit]:
            padded = r + [""] * (len(headers) - len(r))
            lines.append("| " + " | ".join(padded[:len(headers)]) + " |")
        if len(rows) > limit:
            lines.append(f"... ({len(rows) - limit} more rows omitted)")
        return "\n".join(lines)
    except PlaywrightError as e:
        return f"Table Extraction Error: {e}"


# =====================================================================
# 4. Testing & assertions
# =====================================================================

def browser_assert_element(selector: str, state: str = "visible") -> str:
    p = get_page()
    try:
        p.wait_for_selector(selector, state=state, timeout=5000)
        return f"PASS: Element '{selector}' is {state}."
    except PlaywrightError:
        return f"FAIL: Element '{selector}' is NOT {state}."


def browser_assert_text(expected_text: str, selector: str = "body") -> str:
    p = get_page()
    try:
        content = p.inner_text(selector or "body", timeout=5000)
        if expected_text in content:
            return f"PASS: Found '{expected_text}'."
        return f"FAIL: '{expected_text}' not found."
    except PlaywrightError as e:
        return f"Assertion Error: {e}"


def browser_get_console_errors() -> str:
    errors = _state.get("console_errors", [])
    failed = _state.get("failed_requests", [])
    report = []
    if errors:
        report.append("Console Errors:\n" + "\n".join(f"- {e}" for e in errors[-10:]))
    if failed:
        report.append("Failed Requests:\n" + "\n".join(f"- {f}" for f in failed[-10:]))
    return "\n\n".join(report) if report else "PASS: No console errors or failed requests recorded."


# =====================================================================
# 5. Screenshots & PDF (always inside the workspace)
# =====================================================================

def browser_screenshot(path: str = "reports/screenshot.png", selector: str = "") -> str:
    p = get_page()
    target = safe(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        if selector and selector.strip():
            el = p.query_selector(selector.strip())
            if not el:
                return f"Selector '{selector}' not found."
            el.screenshot(path=str(target))
        else:
            p.screenshot(path=str(target), full_page=True)
        return f"Screenshot saved: {target.relative_to(ROOT)}"
    except PlaywrightError as e:
        return f"Screenshot Error: {e}"


def browser_export_pdf(path: str = "reports/webpage.pdf") -> str:
    p = get_page()
    target = safe(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        p.pdf(path=str(target), print_background=True)
        return f"PDF saved: {target.relative_to(ROOT)}"
    except PlaywrightError:
        # Chromium only prints PDFs in headless mode -> use a temporary headless browser
        try:
            hl = _state["pw"].chromium.launch(headless=True)
            try:
                pg = hl.new_page()
                pg.goto(p.url, wait_until="networkidle", timeout=15000)
                pg.pdf(path=str(target), print_background=True)
            finally:
                hl.close()
            return f"PDF saved: {target.relative_to(ROOT)}"
        except Exception as e:
            return f"PDF export error: {e}"


# =====================================================================
# 6. Tool registry (OpenAI/Groq function schemas)
# =====================================================================
# All parameters are declared as strings on purpose: models often send numbers
# and booleans as strings, and strict schema validation would reject them.
# The functions above convert them safely.

def _schema(name: str, desc: str, props: Dict[str, str], required: List[str]) -> Dict[str, Any]:
    return {
        "type": "function",
        "function": {
            "name": name,
            "description": desc,
            "parameters": {
                "type": "object",
                "properties": {k: {"type": "string", "description": d} for k, d in props.items()},
                "required": required,
            },
        },
    }


BROWSER_SCHEMAS = [
    _schema("search_web", "Search Google and return the result text",
            {"query": "Search query"}, ["query"]),
    _schema("browser_open", "Open a website URL in the browser",
            {"url": "Web address"}, ["url"]),
    _schema("browser_navigate", "Navigate history: 'back', 'forward' or 'reload'",
            {"action": "back | forward | reload"}, ["action"]),
    _schema("browser_wait", "Wait for a selector to appear, or pause",
            {"selector": "CSS selector (optional)", "timeout_ms": "Milliseconds, e.g. 2000"}, []),
    _schema("browser_click", "Click an element by CSS selector",
            {"selector": "CSS selector"}, ["selector"]),
    _schema("browser_fill", "Type text into an input by CSS selector",
            {"selector": "CSS selector", "text": "Text to enter",
             "clear_first": "'true' to replace existing text, 'false' to append"},
            ["selector", "text"]),
    _schema("browser_select", "Select an option in a <select> dropdown by value",
            {"selector": "CSS selector", "value": "Option value"}, ["selector", "value"]),
    _schema("browser_check", "Check or uncheck a checkbox/radio",
            {"selector": "CSS selector", "checked": "'true' or 'false'"}, ["selector"]),
    _schema("browser_press_key", "Press a keyboard key like Enter, Tab, Escape",
            {"key": "Key name"}, ["key"]),
    _schema("browser_scroll", "Scroll the page up or down",
            {"direction": "up | down", "amount": "Pixels, e.g. 600"}, []),
    _schema("browser_hover", "Hover the mouse over an element",
            {"selector": "CSS selector"}, ["selector"]),
    _schema("browser_extract_text", "Get visible text of the page or an element",
            {"selector": "CSS selector (default body)", "max_length": "Max characters"}, []),
    _schema("browser_extract_links", "Extract page links, optionally filtered by keyword",
            {"filter_keyword": "Keyword to filter links", "limit": "Max links"}, []),
    _schema("browser_extract_table", "Extract an HTML table as a markdown table",
            {"selector": "CSS selector (default table)", "max_rows": "Max rows"}, []),
    _schema("browser_assert_element", "Assert an element is 'visible', 'hidden' or 'attached'",
            {"selector": "CSS selector", "state": "visible | hidden | attached"}, ["selector"]),
    _schema("browser_assert_text", "Assert specific text is present on the page",
            {"expected_text": "Text to find", "selector": "CSS selector (default body)"},
            ["expected_text"]),
    _schema("browser_get_console_errors", "Get recorded JS console errors and failed requests", {}, []),
    _schema("browser_screenshot", "Capture a screenshot of the page or an element into the workspace",
            {"path": "e.g. reports/shot.png", "selector": "Optional element selector"}, []),
    _schema("browser_export_pdf", "Export the current page as a PDF into the workspace",
            {"path": "e.g. reports/page.pdf"}, []),
]

BROWSER_FUNCS = {
    "search_web": search_web,
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