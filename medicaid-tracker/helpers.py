"""Shared building blocks: downloading files, driving a browser, and reading numbers out of files."""
import base64
import re
from datetime import date
from pathlib import Path

import requests

HERE = Path(__file__).parent
DOWNLOADS = HERE / "downloads"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/130.0 Safari/537.36"
)
MONTHS = ["january", "february", "march", "april", "may", "june",
          "july", "august", "september", "october", "november", "december"]


# ---------- months ----------

def month_key(year, month):
    """2026, 8 -> '2026-08'. All data is keyed by month this way."""
    return f"{int(year):04d}-{int(month):02d}"


def month_number(name):
    """'Aug', 'AUGUST', 'Sept.' -> 8."""
    name = name.strip().lower().rstrip(".")[:3]
    return [m[:3] for m in MONTHS].index(name) + 1


def add_months(key, n):
    y, m = map(int, key.split("-"))
    total = y * 12 + (m - 1) + n
    return month_key(total // 12, total % 12 + 1)


def to_int(text):
    """'1,234,567' -> 1234567; returns None for blanks or non-numbers."""
    try:
        return int(round(float(str(text).replace(",", "").strip())))
    except (ValueError, TypeError):
        return None


def clean(text):
    """Collapse whitespace and drop invisible characters some state sites sprinkle into link text."""
    return " ".join(str(text).replace("​", "").split())


# ---------- plain downloads (sites that don't block scripts) ----------

def http_get(url, timeout=120):
    r = requests.get(url, headers={"User-Agent": USER_AGENT}, timeout=timeout)
    r.raise_for_status()
    return r


def cached_download(state, filename, url=None, fetch=None, refresh=False):
    """Save a file under downloads/<state>/ and reuse it next time.

    refresh=True re-downloads (for files the state overwrites in place, like a running history file).
    Pass either a url (plain download) or fetch, a function returning bytes (browser download).
    """
    path = DOWNLOADS / state / filename
    if path.exists() and not refresh:
        return path
    path.parent.mkdir(parents=True, exist_ok=True)
    data = fetch() if fetch else http_get(url).content
    path.write_bytes(data)
    return path


# ---------- browser (sites that block scripts) ----------

class Browser:
    """A visible Chrome window. Some state sites refuse hidden/automated browsers, so this one is visible.

    Usage:
        with Browser() as b:
            links = b.links("https://...")
            data = b.fetch_bytes("https://.../file.pdf")
    """

    def __enter__(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self._browser = self._pw.chromium.launch(headless=False)
        self.page = self._browser.new_page(viewport={"width": 1280, "height": 720})
        return self

    def __exit__(self, *exc):
        self._browser.close()
        self._pw.stop()

    def open(self, url, wait_ms=3000):
        self.page.goto(url, wait_until="domcontentloaded", timeout=90000)
        self.page.wait_for_timeout(wait_ms)

    def links(self, url=None, wait_ms=3000):
        """All links on the page as (text, absolute url) pairs."""
        if url:
            self.open(url, wait_ms)
        pairs = self.page.eval_on_selector_all(
            "a[href]", "els => els.map(e => [e.innerText, e.href])"
        )
        return [(clean(t), h) for t, h in pairs]

    def fetch_bytes(self, url):
        """Download a file from inside the page, so the request carries the browser's cookies."""
        result = self.page.evaluate(
            """async (u) => {
                const r = await fetch(u, {credentials: 'include'});
                const buf = new Uint8Array(await r.arrayBuffer());
                let s = '';
                for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode(...buf.subarray(i, i + 0x8000));
                return {status: r.status, b64: btoa(s)};
            }""",
            url,
        )
        if result["status"] != 200:
            raise RuntimeError(f"HTTP {result['status']} for {url}")
        return base64.b64decode(result["b64"])


# ---------- reading files ----------

def pdf_lines(path, pages=None):
    """All text lines from a PDF (optionally only some page indexes)."""
    import pdfplumber
    lines = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            if pages is None or i in pages:
                lines += (page.extract_text() or "").splitlines()
    return lines


def numbers_in(line):
    """All comma-formatted or plain integers in a line of text."""
    return [to_int(n) for n in re.findall(r"\d[\d,]*", line)]


def today_key():
    t = date.today()
    return month_key(t.year, t.month)
