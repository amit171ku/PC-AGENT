import base64
import functools
import ipaddress
import os
import re
import socket
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional
from urllib.parse import parse_qs, quote_plus, urlparse

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
    "snap": None,        # last snapshot: its URL + what each number pointed to
    "note": "",          # one-time notice for the agent (e.g. temporary profile in use)
    "crashed": False,
}

# Confirmation callback: fn(prompt_text) -> bool
_confirm_callback: Optional[Callable[[str], bool]] = None

SENSITIVE_WORDS = ("password", "passwd", "pwd", "card", "cvv", "cvc", "secret", "ssn")
MAX_LOG_ITEMS = 100


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


# ---------- URL safety ----------
def _allow_private() -> bool:
    return os.getenv("BROWSER_ALLOW_PRIVATE", "").strip().lower() in ("1", "true", "yes", "on")


def _blocked_reason(url: str) -> str:
    """Return a reason string if the URL must not be opened, or '' if it is fine.

    Blocks non-http(s) schemes and hosts that resolve to local/private network
    addresses (localhost, 192.168.x.x, 10.x.x.x, 169.254.x.x cloud metadata, ...).
    Set BROWSER_ALLOW_PRIVATE=1 in .env to allow private addresses.
    """
    try:
        parts = urlparse(url)
    except ValueError:
        return "Blocked: invalid URL."
    if parts.scheme not in ("http", "https"):
        return f"Blocked: scheme '{parts.scheme}' is not allowed (only http/https)."
    host = parts.hostname
    if not host:
        return "Blocked: URL has no host."
    if _allow_private():
        return ""
    try:
        infos = socket.getaddrinfo(host, None)
    except socket.gaierror:
        return ""  # DNS failure: the browser will report it itself
    for info in infos:
        try:
            ip = ipaddress.ip_address(info[4][0].split("%")[0])
        except ValueError:
            continue
        if ip.version == 6 and ip.ipv4_mapped:
            ip = ip.ipv4_mapped
        if (ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved
                or ip.is_multicast or ip.is_unspecified):
            return f"Blocked: '{host}' points to a local/private network address ({ip})."
    return ""


def _route_guard(route, request):
    """Checks every page navigation (including redirects) against _blocked_reason."""
    reason = ""
    try:
        if request.is_navigation_request():
            reason = _blocked_reason(request.url)
    except Exception:
        reason = "Blocked: URL check failed."
    if reason:
        print(f"[browser] {reason} ({request.url[:120]})")
        route.abort()
    else:
        route.continue_()


def _field_problem(p, selector: str) -> str:
    """Inspect the real element (not just the selector text). Returns a reason if sensitive."""
    try:
        info = p.locator(selector).first.evaluate(
            """el => ({
                type: (el.getAttribute('type') || '').toLowerCase(),
                autocomplete: (el.getAttribute('autocomplete') || '').toLowerCase(),
                attrs: [el.name, el.id, el.placeholder, el.getAttribute('aria-label')]
                        .join(' ').toLowerCase()
            })""",
            timeout=5000,
        )
    except PlaywrightError:
        return ""  # element missing: the fill itself will report the error
    if info.get("type") == "password":
        return "password field"
    ac = info.get("autocomplete", "")
    if ac.startswith("cc-") or "one-time-code" in ac or "password" in ac:
        return f"sensitive autocomplete field ({ac})"
    if any(w in info.get("attrs", "") for w in SENSITIVE_WORDS):
        return "sensitive field name"
    return ""


# ---------- Output paths ----------
def _unique_path(target: Path) -> Path:
    """Never overwrite an existing file: add a timestamp if the name is taken."""
    if not target.exists():
        return target
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    cand = target.with_name(f"{target.stem}_{stamp}{target.suffix}")
    n = 1
    while cand.exists():
        cand = target.with_name(f"{target.stem}_{stamp}_{n}{target.suffix}")
        n += 1
    return cand


# ---------- Browser lifecycle ----------
def _on_console(msg):
    if msg.type == "error":
        errs = _state["console_errors"]
        errs.append(msg.text)
        del errs[:-MAX_LOG_ITEMS]


def _on_request_failed(req):
    fails = _state["failed_requests"]
    fails.append(f"{req.method} {req.url} - {req.failure}")
    del fails[:-MAX_LOG_ITEMS]


_DEAD_MARKERS = ("has been closed", "target closed", "crashed", "browser closed",
                 "connection closed", "target page, context or browser")


def _is_dead_text(text: str) -> bool:
    t = (text or "").lower()
    return any(m in t for m in _DEAD_MARKERS)


def _short_err(e) -> str:
    """First meaningful line of a Playwright error (the full call log is too long for the agent)."""
    lines = [ln.strip() for ln in str(e).splitlines() if ln.strip()]
    return (lines[0] if lines else "unknown error")[:200]


def _persistent() -> bool:
    """Keep cookies/logins between runs (default ON). BROWSER_PERSISTENT=0 gives a fresh browser each run."""
    return os.getenv("BROWSER_PERSISTENT", "1").strip().lower() not in ("0", "false", "no", "off")


def _profile_dir() -> Path:
    raw = os.getenv("BROWSER_PROFILE_DIR", "").strip()
    return Path(raw) if raw else ROOT / ".browser_profile"


def _on_crash(pg):
    print("[browser] page crashed: the browser will be restarted on the next step")
    _state["crashed"] = True


def _attach_page(pg):
    pg.on("console", _on_console)
    pg.on("requestfailed", _on_request_failed)
    pg.on("crash", _on_crash)


def _on_new_page(pg):
    """A click opened a new tab/popup: make it the active page."""
    _attach_page(pg)
    _state["page"] = pg


def _alive(pg) -> bool:
    """Cheap health check: a closed or crashed page cannot run JavaScript."""
    try:
        if pg.is_closed():
            return False
        pg.evaluate("1")
        return True
    except PlaywrightError as e:
        return not _is_dead_text(str(e))   # e.g. a navigation in progress is NOT a crash
    except Exception:
        return False


def _pid_alive(pid: int) -> bool:
    if os.name != "posix":
        return True          # never guess on Windows (os.kill would terminate the process)
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except Exception:
        return True
    return True


def _clear_stale_lock(profile: Path):
    """Linux/macOS only: remove Chromium's lock files when the process that owned them is gone."""
    if os.name != "posix":
        return
    try:
        lock = profile / "SingletonLock"
        if lock.is_symlink():
            if _pid_alive(int(os.readlink(lock).rsplit("-", 1)[-1])):
                return
        elif not lock.exists():
            return
        for name in ("SingletonLock", "SingletonCookie", "SingletonSocket"):
            try:
                (profile / name).unlink()
            except OSError:
                pass
        print("[browser] removed a stale profile lock")
    except Exception:
        pass


_LOCK_NOTE = ("NOTICE: the saved browser profile is in use by another Chrome window or agent, so this run "
              "uses a temporary profile (no saved logins). Close the other window and restart the agent "
              "to use saved logins again.")


def _launch_context(pw, args, ua, viewport):
    """Open the persistent profile. If it is locked, retry once, then fall back to a temporary profile."""
    profile = _profile_dir()
    profile.mkdir(parents=True, exist_ok=True)
    for attempt in (1, 2):
        try:
            ctx = pw.chromium.launch_persistent_context(
                str(profile), headless=False, args=args, viewport=viewport, user_agent=ua)
            return None, ctx
        except Exception as e:
            print(f"[browser] profile launch failed (try {attempt}/2): {_short_err(e)}")
            _clear_stale_lock(profile)
            time.sleep(1.5)
    print("[browser] profile looks locked by another browser: using a temporary profile this run")
    browser = pw.chromium.launch(headless=False, args=args)
    ctx = browser.new_context(viewport=viewport, user_agent=ua)
    _state["note"] = _LOCK_NOTE
    return browser, ctx


def get_page():
    """Lazily start Chromium and return the active Page. Restarts it if the window was closed or crashed."""
    pg = _state.get("page")
    ctx = _state.get("context")
    br = _state.get("browser")          # None when using a persistent profile
    try:
        if ctx is not None and not _state.get("crashed") and (br is None or br.is_connected()):
            if pg is None or pg.is_closed():
                open_pages = [x for x in ctx.pages if not x.is_closed()]
                pg = open_pages[-1] if open_pages else None
                _state["page"] = pg
            if pg is not None and _alive(pg) and not _state.get("crashed"):
                return pg
    except Exception:
        pass

    if ctx is not None:
        print("[browser] window closed or crashed: starting a fresh browser")
    close_browser()

    pw = sync_playwright().start()
    try:
        args = [
            "--disable-blink-features=AutomationControlled",
            "--disable-infobars",
            "--start-maximized",
        ]
        # The Chromium sandbox is a security feature. Only disable it if you really must
        # (for example when running as root in Docker): set BROWSER_NO_SANDBOX=1 in .env.
        if os.getenv("BROWSER_NO_SANDBOX", "").strip().lower() in ("1", "true", "yes", "on"):
            args.append("--no-sandbox")
        ua = (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
        )
        viewport = {"width": 1280, "height": 800}

        if _persistent():
            browser, context = _launch_context(pw, args, ua, viewport)
        else:
            browser = pw.chromium.launch(headless=False, args=args)
            context = browser.new_context(viewport=viewport, user_agent=ua)
        page = context.pages[0] if context.pages else context.new_page()

        context.route("**/*", _route_guard)   # block navigation to local/private addresses
        context.add_init_script(
            "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"
        )
        _attach_page(page)
        context.on("page", _on_new_page)
    except Exception:
        try:
            pw.stop()
        except Exception:
            pass
        raise

    _state.update({"pw": pw, "browser": browser, "context": context, "page": page})
    return page


def close_browser():
    """Close the browser and reset all tracking state."""
    try:
        if _state.get("browser"):
            _state["browser"].close()
        elif _state.get("context"):
            _state["context"].close()
    except Exception:
        pass
    try:
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
        "snap": None,
        "note": "",
        "crashed": False,
    })


def pump(ms: int = 100):
    """Keep the browser alive while Python is waiting (for you, for a new chat message, ...).

    Playwright's sync API only serves browser events (page loads, our URL guard, redirects) while Python is
    inside a Playwright call. If Python just sleeps, pages you use BY HAND (e.g. logging in) freeze and
    look like they are buffering forever. Every wait loop must call this.
    """
    try:
        pg = _state.get("page")
        if pg is None or pg.is_closed():
            ctx = _state.get("context")
            open_pages = [x for x in ctx.pages if not x.is_closed()] if ctx is not None else []
            pg = open_pages[-1] if open_pages else None
        if pg is not None:
            pg.wait_for_timeout(ms)
            return
    except Exception:
        pass
    time.sleep(ms / 1000)


def _take_note() -> str:
    """One-time notice for the agent (shown once, then cleared)."""
    n = _state.get("note", "")
    _state["note"] = ""
    return n


_ERR_START = re.compile(r"^[A-Za-z ]{0,30}(error|failed)", re.I)


def _resilient(fn):
    """If a tool fails because the browser was closed or crashed, reset it and tell the agent."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            result = fn(*args, **kwargs)
        except PlaywrightError as e:
            result = f"Browser Error: {_short_err(e)}"
        if isinstance(result, str) and _ERR_START.match(result) and _is_dead_text(result):
            close_browser()
            return (result + "\nThe browser was closed or crashed and has been reset. "
                    "Call browser_open again to continue (old snapshot numbers are gone).")
        return result
    return wrapper


# =====================================================================
# 1. Search & Browsing
# =====================================================================

# --- Structured search results (title + URL + snippet) so the agent can cite sources ---
def _unwrap_redirect(url: str) -> str:
    """Bing/Google wrap result links in redirects. Return the real destination when it can be decoded."""
    try:
        parts = urlparse(url)
        host = (parts.hostname or "").lower()
        if host.endswith("bing.com") and parts.path.startswith("/ck/"):
            u = parse_qs(parts.query).get("u", [""])[0]
            if u.startswith("a1"):                      # "a1" + base64url(real URL)
                raw = u[2:]
                raw += "=" * (-len(raw) % 4)
                real = base64.urlsafe_b64decode(raw).decode("utf-8", "ignore")
                if real.startswith(("http://", "https://")):
                    return real
        if host.endswith("google.com") and parts.path == "/url":
            q = parse_qs(parts.query).get("q", [""])[0]
            if q.startswith(("http://", "https://")):
                return q
    except Exception:
        pass
    return url


def _result_title(r, url: str) -> str:
    """Never leave a result without a title: title -> site line -> host name."""
    t = re.sub(r"\s+", " ", (r.get("title") or "")).strip()
    if t:
        return t[:120]
    site = re.sub(r"\s+", " ", (r.get("site") or "")).strip()
    if site:
        return site[:120]
    return (urlparse(url).hostname or "untitled result")[:120]


def _dedupe(items):
    """Normalise raw results: real URLs, a title for every result, no duplicates, no search-engine junk."""
    seen, out = set(), []
    for r in items or []:
        url = _unwrap_redirect((r.get("url") or "").strip())
        if not url.startswith(("http://", "https://")) or url in seen:
            continue
        parts = urlparse(url)
        host = (parts.hostname or "").lower()
        if host.endswith("bing.com") and not parts.path.startswith("/ck/"):
            continue                                    # related searches, image tabs, ...
        seen.add(url)
        out.append({"title": _result_title(r, url), "url": url,
                    "snippet": re.sub(r"\s+", " ", r.get("snippet") or "").strip()})
    return out


def _fmt_results(items, limit: int = 8) -> str:
    """Format normalised results as a numbered list that always includes the URL."""
    out = []
    for i, r in enumerate(items[:limit], 1):
        out.append(f"{i}. {r['title']}\n   URL: {r['url']}\n   {r['snippet'][:250]}")
    return "\n".join(out)


_GOOGLE_JS = r"""els => els.map(a => {
    const h = a.querySelector('h3');
    const box = a.closest('div.g') || (a.parentElement && a.parentElement.parentElement);
    return {title: h ? h.innerText.trim() : '', url: a.href,
            snippet: box ? box.innerText.replace(/\s+/g, ' ') : ''};
}).filter(x => x.url.startsWith('http') && !x.url.includes('google.com/search'))"""

_BING_JS = r"""els => els.map(li => {
    const a = li.querySelector('h2 a') || li.querySelector('a[href^="http"]');
    const h = li.querySelector('h2');
    const p = li.querySelector('.b_caption p') || li.querySelector('p');
    const cite = li.querySelector('cite');
    const title = a ? (a.innerText || a.getAttribute('aria-label') || a.title || (h ? h.innerText : '') || '') : '';
    return {title: title.trim(), url: a ? a.href : '',
            snippet: p ? p.innerText.trim() : '', site: cite ? cite.innerText.trim() : ''};
}).filter(x => x.url && x.url.startsWith('http'))"""


def search_web(query: str) -> str:
    """Search Google (falls back to Bing if blocked). Returns results WITH URLs."""
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
        # "robot" alone is NOT used: normal results about robots would trigger the fallback.
        blocked = any(
            phrase in body_text.lower()
            for phrase in ("unusual traffic", "recaptcha", "solve the challenge",
                           "not a robot", "bot detection")
        )

        if not blocked:
            try:
                p.wait_for_selector("#search", timeout=3000)
                # 1) Structured results with URLs (preferred)
                try:
                    items = _dedupe(p.eval_on_selector_all("#search a:has(h3)", _GOOGLE_JS))
                    if items:
                        return _fmt_results(items)
                except Exception:
                    pass
                # 2) Plain text fallback (no URLs)
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
            p.wait_for_selector("#b_results", timeout=4000)
            # Structured results with URLs (preferred). Note: Bing may return
            # redirect links (bing.com/ck/a...); browser_open follows them.
            try:
                items = _dedupe(p.eval_on_selector_all("#b_results li.b_algo", _BING_JS))
                if items:
                    return _fmt_results(items)
            except Exception:
                pass
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
    url = (url or "").strip()
    if not url.startswith(("http://", "https://")):
        url = "https://" + url
    reason = _blocked_reason(url)
    if reason:
        return reason
    try:
        p = get_page()
        resp = p.goto(url, wait_until="domcontentloaded", timeout=30000)
        status = resp.status if resp else "unknown"
        note = "\n".join(x for x in (_page_notice(p), _take_note()) if x)
        return (f"Opened: {p.title()} (Status: {status}) - {p.url}"
                + (f"\n{note}" if note else ""))
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
    blocked = _action_guard(p)
    if blocked:
        return blocked
    if not _require_confirm(_card(("Action", "Click"), ("Element", selector), ("Site", _site(p)))):
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
    blocked = _action_guard(p)
    if blocked:
        return blocked
    problem = _field_problem(p, selector)
    if problem:
        return f"Refused: safety filter prevents typing into a {problem}."
    if not _require_confirm(_card(("Action", "Type text"), ("Field", selector),
                                  ("Value", text[:200]), ("Site", _site(p)))):
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
    blocked = _action_guard(p)
    if blocked:
        return blocked
    if not _require_confirm(_card(("Action", "Select option"), ("Field", selector),
                                  ("Value", value), ("Site", _site(p)))):
        return "User denied."
    try:
        p.select_option(selector, value=value, timeout=10000)
        return f"Selected '{value}' in '{selector}'."
    except PlaywrightError as e:
        return f"Select Error: {e}"


def browser_check(selector: str, checked: str = "true") -> str:
    want = _bool(checked)
    p = get_page()
    blocked = _action_guard(p)
    if blocked:
        return blocked
    if not _require_confirm(_card(("Action", "Check" if want else "Uncheck"),
                                  ("Element", selector), ("Site", _site(p)))):
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
    blocked = _action_guard(p)
    if blocked:
        return blocked
    if not _require_confirm(_card(("Action", "Press key"), ("Key", key), ("Site", _site(p)))):
        return "User denied."
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
        # Always carry the source so the agent can cite it
        header = f"Source: {p.title()} - {p.url}\n"
        text = p.inner_text(selector or "body", timeout=10000).strip()
        if len(text) > limit:
            return header + text[:limit] + f"\n... [truncated, total {len(text)} chars]"
        return header + (text or "(empty)")
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
# 3b. Working on logged-in websites (snapshot + click/fill by number)
# =====================================================================

def _allowed_domains() -> List[str]:
    raw = os.getenv("BROWSER_ALLOWED_DOMAINS", "")
    return [d.strip().lower().lstrip(".") for d in raw.split(",") if d.strip()]


def _action_guard(p) -> str:
    """If BROWSER_ALLOWED_DOMAINS is set, clicks/typing are only allowed on those sites."""
    domains = _allowed_domains()
    if not domains:
        return ""
    host = (urlparse(p.url).hostname or "").lower()
    if any(host == d or host.endswith("." + d) for d in domains):
        return ""
    return (f"Blocked: actions are only allowed on {', '.join(domains)} "
            f"(current page: {host or 'blank'}). Ask the user to add this site to "
            f"BROWSER_ALLOWED_DOMAINS in .env if it is intended.")


_BLOCK_PHRASES = ("verify you are human", "unusual traffic", "not a robot",
                  "security check", "checking your browser", "are you a human")


def _page_notice(p) -> str:
    """Detect CAPTCHA / login pages so the agent stops and asks the user instead of pushing on."""
    notes = []
    try:
        text = p.inner_text("body", timeout=3000).lower()[:6000]
    except Exception:
        text = ""
    try:
        captcha_frame = p.locator(
            "iframe[src*='recaptcha'], iframe[src*='hcaptcha'], iframe[src*='captcha']").count() > 0
    except Exception:
        captcha_frame = False
    if captcha_frame or any(ph in text for ph in _BLOCK_PHRASES):
        notes.append("NOTICE: CAPTCHA / human verification detected. STOP. Use ask_user to tell the "
                     "user to solve it in the browser window, then continue after they reply.")
    try:
        pw_field = p.locator("input[type='password']").first
        if pw_field.count() > 0 and pw_field.is_visible():
            notes.append("NOTICE: Login page detected. STOP. Use ask_user to tell the user to log in "
                         "manually in the browser window (never type passwords), then call "
                         "browser_snapshot after they reply.")
    except Exception:
        pass
    return "\n".join(notes)


_SNAPSHOT_JS = r"""(limit) => {
  document.querySelectorAll('[data-agent-ref]').forEach(e => {
    e.removeAttribute('data-agent-ref');
    e.removeAttribute('data-agent-label');
  });
  const sel = 'a[href], button, input:not([type=hidden]), textarea, select, [role=button], [role=link], [role=checkbox], [role=tab], [role=menuitem], [contenteditable]';
  const fields = 'input:not([type=hidden]), textarea, select';
  const clean = s => (s || '').replace(/\s+/g, ' ').trim();
  const byIds = ids => (ids || '').split(/\s+/).map(id => {
    const n = id ? document.getElementById(id) : null;
    return n ? clean(n.innerText || n.textContent) : '';
  }).filter(Boolean).join(' ');
  const generic = s => !s || /^(your answer|type here|enter text|enter|text|answer|search|input)$/i.test(s);
  // nearest question heading above a field (e.g. the question text in a Google Form)
  const groupLabel = el => {
    let node = el.parentElement;
    for (let d = 0; node && d < 7; d++, node = node.parentElement) {
      if (node.querySelectorAll(fields).length > 8) break;
      const h = node.querySelector('legend, h1, h2, h3, h4, h5, h6, [role=heading]');
      if (h && !h.contains(el)) {
        const t = clean(h.innerText);
        if (t) return t;
      }
    }
    return '';
  };
  const out = [];
  for (const el of document.querySelectorAll(sel)) {
    if (out.length >= limit) break;
    const r = el.getBoundingClientRect();
    const st = getComputedStyle(el);
    if (r.width < 2 || r.height < 2 || st.visibility === 'hidden' || st.display === 'none' || el.disabled) continue;
    const tag = el.tagName.toLowerCase();
    const type = (el.getAttribute('type') || '').toLowerCase();
    const lab = (el.labels && el.labels[0]) ? el.labels[0].innerText : '';
    let label = clean(byIds(el.getAttribute('aria-labelledby')) || el.getAttribute('aria-label') || lab ||
                      el.innerText || el.placeholder || el.getAttribute('title') || el.name || el.id);
    const isField = tag === 'input' || tag === 'textarea' || tag === 'select' || el.isContentEditable;
    if (isField && generic(label)) {
      const g = groupLabel(el);
      if (g) label = g;
    }
    label = label.slice(0, 80);
    let value = '';
    if (type === 'password') value = '(hidden)';
    else if (tag === 'input' || tag === 'textarea' || tag === 'select') value = clean(el.value).slice(0, 80);
    const ref = out.length + 1;
    el.setAttribute('data-agent-ref', String(ref));
    el.setAttribute('data-agent-label', label);
    out.push({ref: ref, tag: tag, type: type, label: label, value: value,
              checked: !!el.checked, href: tag === 'a' ? el.href.slice(0, 120) : ''});
  }
  return out;
}"""


def _same_page(a: str, b: str) -> bool:
    return (a or "").split("#")[0] == (b or "").split("#")[0]


def _take_snapshot(p, limit: int = 60):
    """Run the snapshot script and remember what each number pointed to (for stale-number recovery)."""
    try:
        p.wait_for_load_state("networkidle", timeout=3000)
    except PlaywrightError:
        pass
    items = p.evaluate(_SNAPSHOT_JS, max(5, min(limit, 150)))
    _state["snap"] = {"url": p.url, "refs": {it["ref"]: (it["tag"], it["label"]) for it in items}}
    return items


def _format_snapshot(p, items) -> str:
    lines = [f"Page: {p.title()} - {p.url}"]
    note = _page_notice(p)
    if note:
        lines.append(note)
    if not items:
        lines.append("(no interactive elements found; try browser_extract_text or browser_scroll)")
    for it in items:
        kind = it["tag"] + (f"({it['type']})" if it["type"] else "")
        row = f"[{it['ref']}] {kind} \"{it['label']}\""
        if it["value"]:
            row += f" = \"{it['value']}\""
        if it["type"] in ("checkbox", "radio"):
            row += " (checked)" if it["checked"] else " (unchecked)"
        if it["href"]:
            row += f" -> {it['href']}"
        lines.append(row)
    lines.append("Use browser_click_ref / browser_fill_ref with these numbers. Numbers go stale after "
                 "any click or page change: call browser_snapshot again. For dropdowns/checkboxes use "
                 "browser_select / browser_check with selector [data-agent-ref=\"N\"].")
    return "\n".join(lines)


def browser_snapshot(limit: str = "60") -> str:
    """List visible interactive elements as numbered refs the agent can click or fill."""
    p = get_page()
    try:
        text = _format_snapshot(p, _take_snapshot(p, _int(limit, 60)))
        note = _take_note()
        return (note + "\n" + text) if note else text
    except PlaywrightError as e:
        return f"Snapshot Error: {_short_err(e)}"


def _stale_reply(p, why: str, items=None) -> str:
    """Tell the agent its number is out of date AND hand it a fresh snapshot (saves a step)."""
    try:
        if items is None:
            items = _take_snapshot(p)
        return f"{why} Fresh snapshot below: choose the new number.\n" + _format_snapshot(p, items)
    except PlaywrightError:
        return f"{why} Call browser_snapshot again."


def _resolve_ref(p, ref):
    """Return (selector, error). Recovers when the page re-rendered: same page + same element label."""
    n = _int(ref, -1)
    if n < 1:
        return "", "Invalid ref. Use a number from browser_snapshot."
    snap = _state.get("snap")
    if not snap:
        return "", "No snapshot yet. Call browser_snapshot first."
    sel = f'[data-agent-ref="{n}"]'
    same = _same_page(p.url, snap["url"])
    if same and p.locator(sel).count() == 1:
        return sel, ""
    old = snap["refs"].get(n)
    items = None
    if same and old:
        try:
            items = _take_snapshot(p)
        except PlaywrightError:
            items = None
        if items is not None:
            matches = [it for it in items if (it["tag"], it["label"]) == old]
            if len(matches) == 1:       # the same element, just renumbered after a re-render
                return f'[data-agent-ref="{matches[0]["ref"]}"]', ""
    why = ("The page changed since the last snapshot, so that number is out of date."
           if same else "The browser is on a different page now, so that number is out of date.")
    return "", _stale_reply(p, why, items)


def _click_hint(e) -> str:
    m = str(e).lower()
    if "intercepts pointer events" in m:
        return " A popup or overlay is covering it: close it (or press Escape), then browser_snapshot."
    if "not visible" in m or "outside of the viewport" in m:
        return " It is not visible: use browser_scroll, then browser_snapshot."
    if "timeout" in m:
        return " Take a fresh browser_snapshot: the page may have changed."
    return ""


def _site(p) -> str:
    return (urlparse(p.url).hostname or p.url)[:60]


def _card(*rows) -> str:
    """Short multi-line text for approval dialogs: 'Key: value' per line."""
    return "\n".join(f"{k}: {v}" for k, v in rows if v)


def _describe(p, sel: str) -> str:
    """Readable element description for approval cards, using the label saved by browser_snapshot."""
    try:
        return p.locator(sel).first.evaluate(
            r"""el => {
                const t = el.tagName.toLowerCase();
                const l = (el.getAttribute('data-agent-label') || el.getAttribute('aria-label') ||
                           el.innerText || el.placeholder || el.name || el.id || '')
                          .replace(/\s+/g, ' ').trim().slice(0, 80);
                return t + ' "' + l + '"';
            }""",
            timeout=5000,
        )
    except PlaywrightError:
        return sel


_SUBMIT_WORDS = ("save", "submit", "update", "apply", "send", "post", "publish",
                 "delete", "remove", "confirm", "connect")


def browser_click_ref(ref: str) -> str:
    p = get_page()
    blocked = _action_guard(p)
    if blocked:
        return blocked
    try:
        sel, err = _resolve_ref(p, ref)
        if err:
            return err
        desc = _describe(p, sel)
        risky = any(w in desc.lower() for w in _SUBMIT_WORDS)
        prompt = _card(("Action", "CLICK - may SUBMIT or CHANGE data" if risky else "Click"),
                       ("Element", desc), ("Site", _site(p)))
        if not _require_confirm(prompt):
            return "User denied."
        if p.locator(sel).count() != 1:     # the page may have changed while you were deciding
            return _stale_reply(p, "The page changed while waiting for your approval, so nothing was clicked.")
        p.click(sel, timeout=10000)
        p.wait_for_timeout(1000)
        p = get_page()     # a click may have opened a new tab
        note = _page_notice(p)
        return (f"Clicked {desc}. Now on: {p.title()} - {p.url}. Call browser_snapshot to see the page."
                + (f"\n{note}" if note else ""))
    except PlaywrightError as e:
        return f"Click Error: {_short_err(e)}.{_click_hint(e)}"


def browser_fill_ref(ref: str, text: str, clear_first: str = "true") -> str:
    p = get_page()
    blocked = _action_guard(p)
    if blocked:
        return blocked
    try:
        sel, err = _resolve_ref(p, ref)
        if err:
            return err
        problem = _field_problem(p, sel)
        if problem:
            return f"Refused: safety filter prevents typing into a {problem}."
        desc = _describe(p, sel)
        if any(w in desc.lower() for w in SENSITIVE_WORDS):
            return "Refused: safety filter prevents typing into sensitive fields."
        if not _require_confirm(_card(("Action", "Type text"), ("Field", desc),
                                      ("Value", text[:200]), ("Site", _site(p)))):
            return "User denied."
        if p.locator(sel).count() != 1:
            return _stale_reply(p, "The page changed while waiting for your approval, so nothing was typed.")
        if _bool(clear_first):
            p.fill(sel, text, timeout=10000)
        else:
            p.locator(sel).press_sequentially(text, timeout=10000)
        return f"Typed into {desc}."
    except PlaywrightError as e:
        return f"Fill Error: {_short_err(e)}.{_click_hint(e)}"


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
# 5. Screenshots & PDF (always inside the workspace, never overwrite)
# =====================================================================

def browser_screenshot(path: str = "reports/screenshot.png", selector: str = "") -> str:
    p = get_page()
    target = safe(path)
    if target.suffix.lower() not in (".png", ".jpg", ".jpeg"):
        target = target.with_suffix(".png")
    target = _unique_path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        if selector and selector.strip():
            el = p.query_selector(selector.strip())
            if not el:
                return f"Selector '{selector}' not found."
            el.screenshot(path=str(target))
        else:
            p.screenshot(path=str(target), full_page=True)
        return f"Screenshot saved: {target.relative_to(ROOT).as_posix()}"
    except PlaywrightError as e:
        return f"Screenshot Error: {e}"


def browser_export_pdf(path: str = "reports/webpage.pdf") -> str:
    p = get_page()
    target = safe(path)
    if target.suffix.lower() != ".pdf":
        target = target.with_suffix(".pdf")
    target = _unique_path(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    try:
        p.pdf(path=str(target), print_background=True)
        return f"PDF saved: {target.relative_to(ROOT).as_posix()}"
    except PlaywrightError:
        # Chromium only prints PDFs in headless mode -> use a temporary headless browser
        reason = _blocked_reason(p.url)
        if reason:
            return reason
        try:
            hl = _state["pw"].chromium.launch(headless=True)
            try:
                pg = hl.new_page()
                pg.goto(p.url, wait_until="networkidle", timeout=15000)
                pg.pdf(path=str(target), print_background=True)
            finally:
                hl.close()
            return f"PDF saved: {target.relative_to(ROOT).as_posix()}"
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
    _schema("search_web", "Search the web. Returns numbered results with title, URL and snippet. "
            "Open the best URLs with browser_open and read them before answering.",
            {"query": "Search query"}, ["query"]),
    _schema("browser_open", "Open a website URL in the browser (http/https only, no local/private addresses)",
            {"url": "Web address"}, ["url"]),
    _schema("browser_navigate", "Navigate history: 'back', 'forward' or 'reload'",
            {"action": "back | forward | reload"}, ["action"]),
    _schema("browser_wait", "Wait for a selector to appear, or pause",
            {"selector": "CSS selector (optional)", "timeout_ms": "Milliseconds, e.g. 2000"}, []),
    _schema("browser_click", "Click an element by CSS selector",
            {"selector": "CSS selector"}, ["selector"]),
    _schema("browser_fill", "Type text into an input by CSS selector (password/card fields are refused)",
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
    _schema("browser_extract_text", "Get visible text of the page or an element (starts with the page's Source URL)",
            {"selector": "CSS selector (default body)", "max_length": "Max characters"}, []),
    _schema("browser_extract_links", "Extract page links, optionally filtered by keyword",
            {"filter_keyword": "Keyword to filter links", "limit": "Max links"}, []),
    _schema("browser_extract_table", "Extract an HTML table as a markdown table",
            {"selector": "CSS selector (default table)", "max_rows": "Max rows"}, []),
    _schema("browser_snapshot", "List visible clickable/editable elements of the current page as numbered refs [1], [2]... Call before browser_click_ref / browser_fill_ref and again after every click or page change.",
            {"limit": "Max elements, default 60"}, []),
    _schema("browser_click_ref", "Click an element by its browser_snapshot number (asks the user to confirm)",
            {"ref": "Number from browser_snapshot"}, ["ref"]),
    _schema("browser_fill_ref", "Type text into an input/textarea by its browser_snapshot number (password/card fields refused; asks the user to confirm)",
            {"ref": "Number from browser_snapshot", "text": "Text to enter",
             "clear_first": "'true' to replace existing text, 'false' to append"},
            ["ref", "text"]),
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
    "browser_snapshot": browser_snapshot,
    "browser_click_ref": browser_click_ref,
    "browser_fill_ref": browser_fill_ref,
    "browser_assert_element": browser_assert_element,
    "browser_assert_text": browser_assert_text,
    "browser_get_console_errors": browser_get_console_errors,
    "browser_screenshot": browser_screenshot,
    "browser_export_pdf": browser_export_pdf,
}


# Every browser tool now recovers from a closed/crashed browser (see _resilient)
BROWSER_FUNCS = {name: _resilient(fn) for name, fn in BROWSER_FUNCS.items()}