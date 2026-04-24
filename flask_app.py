"""Flask wrapper for EMA RMP Date Finder — used for cloud deployment."""

import json
import os
import re
import urllib.error
import urllib.request
from html.parser import HTMLParser

from flask import Flask, jsonify, request, send_from_directory

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__, static_folder=BASE_DIR, static_url_path="")

EMA_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9",
}


class _TextExtractor(HTMLParser):
    """Extract visible text from HTML, skipping scripts/styles."""

    def __init__(self):
        super().__init__()
        self._parts: list[str] = []
        self._skip = False

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style", "noscript"):
            self._skip = True

    def handle_endtag(self, tag):
        if tag in ("script", "style", "noscript"):
            self._skip = False
        if tag in (
            "div", "p", "h1", "h2", "h3", "h4", "h5", "h6",
            "br", "li", "tr", "td", "th", "dt", "dd",
            "section", "article", "header", "footer",
        ):
            self._parts.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self._parts.append(data)

    def get_text(self) -> str:
        return "".join(self._parts)


def fetch_rmp_date(url: str) -> dict:
    """Fetch an EMA page and extract the RMP date."""
    result = {
        "date_str": None,
        "first_published": None,
        "last_updated": None,
        "error": None,
    }

    if not url or not url.strip():
        result["error"] = "Пустая ссылка"
        return result

    url = url.strip()

    try:
        req = urllib.request.Request(url, headers=EMA_HEADERS)
        with urllib.request.urlopen(req, timeout=30) as resp:
            html = resp.read().decode("utf-8", errors="replace")
    except Exception as exc:
        result["error"] = f"Ошибка загрузки: {exc}"
        return result

    extractor = _TextExtractor()
    try:
        extractor.feed(html)
    except Exception:
        result["error"] = "Ошибка разбора HTML"
        return result

    text = extractor.get_text()
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]

    rmp_idx = None
    for i, line in enumerate(lines):
        if re.search(
            r"EPAR\s*[-\u2013\u2014]\s*Risk[\s\-]+management[\s\-]+plan",
            line,
            re.IGNORECASE,
        ):
            rmp_idx = i
            break

    if rmp_idx is None:
        result["error"] = "RMP-секция не найдена"
        return result

    date_re = re.compile(r"\d{2}/\d{2}/\d{4}")
    section_stop_re = re.compile(
        r"^(Product information|EPAR\s*[-\u2013\u2014]|All Authorised"
        r"|Authorisation details|Product details|Assessment history"
        r"|More information)",
        re.IGNORECASE,
    )

    for line in lines[rmp_idx + 1 : rmp_idx + 15]:
        if line.strip() == "View":
            break
        if section_stop_re.search(line):
            break
        if line.startswith("First published"):
            m = date_re.search(line)
            if m and result["first_published"] is None:
                result["first_published"] = m.group(0)
        elif line.startswith("Last updated"):
            m = date_re.search(line)
            if m and result["last_updated"] is None:
                result["last_updated"] = m.group(0)
        else:
            m = date_re.search(line)
            if m:
                if result["first_published"] is None:
                    result["first_published"] = m.group(0)
                elif result["last_updated"] is None:
                    result["last_updated"] = m.group(0)

    if result["last_updated"]:
        result["date_str"] = result["last_updated"]
    elif result["first_published"]:
        result["date_str"] = result["first_published"]

    if result["date_str"] is None:
        result["error"] = "Даты RMP не найдены"

    return result


@app.route("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/api/fetch-rmp", methods=["POST"])
def api_fetch_rmp():
    data = request.get_json(force=True)
    result = fetch_rmp_date(data.get("url", ""))
    return jsonify(result)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8090))
    app.run(host="0.0.0.0", port=port)
