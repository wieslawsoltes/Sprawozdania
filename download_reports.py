import os
import re
import unicodedata
import urllib.parse
import urllib.request
from html import unescape
from html.parser import HTMLParser
from pathlib import Path

BASE_URL = "https://zopo.bipraciborz.pl"
MAIN_URL = f"{BASE_URL}/bipkod/40495541"
OUTPUT_DIR = Path("pobrane/sprawozdania_2024")


def slugify(text: str) -> str:
    """Convert text to an ASCII safe slug suitable for file and dir names."""
    normalized = unicodedata.normalize("NFKD", text)
    ascii_text = normalized.encode("ascii", "ignore").decode("ascii")
    ascii_text = re.sub(r"[^0-9A-Za-z]+", "_", ascii_text)
    return ascii_text.strip("_") or "plik"


class AnchorParser(HTMLParser):
    """Collect anchor href and text pairs."""

    def __init__(self):
        super().__init__()
        self._in_anchor = False
        self._href = None
        self._text_parts = []
        self.results = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            self._in_anchor = True
            self._href = dict(attrs).get("href")
            self._text_parts = []

    def handle_endtag(self, tag):
        if tag == "a" and self._in_anchor:
            text = unescape("".join(self._text_parts)).strip()
            self.results.append((self._href, text))
            self._in_anchor = False
            self._href = None
            self._text_parts = []

    def handle_data(self, data):
        if self._in_anchor:
            self._text_parts.append(data)


def fetch(url: str) -> str:
    """Fetch URL content as text with a simple User-Agent."""
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def extract_institution_links(html: str):
    """Return list of (name, url) for 2024 institution report pages."""
    parser = AnchorParser()
    parser.feed(html)
    links = []
    prefix = "Sprawozdanie finansowe za rok 2024"
    for href, text in parser.results:
        if href and "bipkod/" in href and prefix in text and href != MAIN_URL:
            name = text.replace(prefix, "").strip()
            links.append((name, urllib.parse.urljoin(BASE_URL, href)))
    return links


def extract_attachment_links(html: str):
    """Return list of (file_title, absolute_url) for attachments on a page."""
    parser = AnchorParser()
    parser.feed(html)
    attachments = []
    for href, text in parser.results:
        if href and "/res/serwisy/pliki/" in href:
            url = urllib.parse.urljoin(BASE_URL, href)
            attachments.append((text or os.path.basename(href), url))
    return attachments


def ensure_unique_path(directory: str, filename: str) -> str:
    """Ensure file path is unique by appending counter when needed."""
    base, ext = os.path.splitext(filename)
    candidate = os.path.join(directory, filename)
    counter = 2
    while os.path.exists(candidate):
        candidate = os.path.join(directory, f"{base}_{counter}{ext}")
        counter += 1
    return candidate


def download_file(url: str, dest_path: str):
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req) as resp, open(dest_path, "wb") as f:
        f.write(resp.read())


def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("Pobieram stronę główną:", MAIN_URL)
    main_html = fetch(MAIN_URL)
    institutions = extract_institution_links(main_html)
    if not institutions:
        raise SystemExit("Nie znaleziono linków do sprawozdań 2024.")

    print(f"Znaleziono {len(institutions)} placówek.")
    for name, url in institutions:
        folder = OUTPUT_DIR / slugify(name)
        print(f"\nPlacówka: {name} -> katalog '{folder}'")
        print("  Pobieram stronę:", url)
        html = fetch(url)
        attachments = extract_attachment_links(html)
        if not attachments:
            print("  Brak załączników na stronie.")
            continue

        for title, file_url in attachments:
            base, ext = os.path.splitext(title)
            ext = ext or ".bin"
            safe_name = f"{slugify(base)}{ext}"
            dest = ensure_unique_path(folder, safe_name)
            print(f"  - zapisuję {safe_name} z {file_url}")
            download_file(file_url, dest)

    print("\\nZakończono pobieranie.")


if __name__ == "__main__":
    main()
