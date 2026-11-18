"""Turn an upload or a public link into source spans with locators."""

from __future__ import annotations

import io
import ipaddress
import re
import socket
from collections.abc import Iterable
from urllib.parse import parse_qs, urlparse

import httpx

class ParseError(Exception):
    pass


AUDIO_EXTENSIONS = {".mp3", ".wav", ".m4a", ".mp4", ".webm", ".mpeg", ".mpga"}
TEXT_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".pptx",
    ".txt",
    ".md",
    ".markdown",
    ".csv",
    ".xlsx",
    ".html",
    ".htm",
    ".srt",
    ".vtt",
}

MAX_UPLOAD_BYTES = 20 * 1024 * 1024
MAX_TEXT_CHARS = 200_000


def extension_of(name: str) -> str:
    base = (name or "").lower().rsplit("/", 1)[-1]
    if "." not in base:
        return ""
    return "." + base.rsplit(".", 1)[-1]


def kind_for(name: str) -> str:
    ext = extension_of(name)
    return {
        ".pdf": "pdf",
        ".docx": "docx",
        ".pptx": "pptx",
        ".xlsx": "xlsx",
        ".csv": "csv",
        ".html": "html",
        ".htm": "html",
        ".srt": "subtitle",
        ".vtt": "subtitle",
        ".md": "md",
        ".markdown": "md",
        ".mp3": "audio",
        ".wav": "audio",
        ".m4a": "audio",
        ".mp4": "video",
        ".webm": "video",
        ".mpeg": "audio",
        ".mpga": "audio",
    }.get(ext, "txt")


def parse_transcript(text: str, locator_prefix: str = "caption") -> list[dict]:
    return _finalize(_plain(text, locator_prefix=locator_prefix))


def parse_upload(name: str, data: bytes) -> list[dict]:
    if len(data) > MAX_UPLOAD_BYTES:
        raise ParseError("Each file needs to be 20 MB or smaller.")
    ext = extension_of(name)
    if ext == ".ppt":
        raise ParseError("Save the presentation as .pptx and upload it again.")
    if ext == ".doc":
        raise ParseError("Save the document as .docx and upload it again.")
    if ext in AUDIO_EXTENSIONS:
        raise ParseError("AUDIO")
    if ext and ext not in TEXT_EXTENSIONS and ext != ".txt":
        raise ParseError(
            "This open edition reads PDF, Word (.docx), PowerPoint (.pptx), "
            "Excel (.xlsx), text, Markdown, HTML, CSV, and subtitles (.srt, .vtt)."
        )
    if ext == ".pdf":
        spans = _pdf(data)
    elif ext == ".docx":
        spans = _docx(data)
    elif ext == ".pptx":
        spans = _pptx(data)
    elif ext == ".xlsx":
        spans = _xlsx(data)
    elif ext in {".html", ".htm"}:
        spans = _html(data.decode("utf-8", errors="replace"), locator_prefix="section")
    elif ext == ".csv":
        spans = _plain(data.decode("utf-8", errors="replace"), locator_prefix="rows")
    else:
        spans = _plain(data.decode("utf-8", errors="replace"), locator_prefix="section")
    return _finalize(spans)


def parse_url(url: str) -> tuple[str, list[dict]]:
    cleaned = (url or "").strip()
    video_id = youtube_id(cleaned)
    if video_id:
        text = _youtube_text(video_id)
        title = f"YouTube {video_id}"
        spans = _plain(text, locator_prefix="caption")
        return title, _finalize(spans)
    _reject_private_url(cleaned)
    title, html = _fetch_html(cleaned)
    spans = _html(html, locator_prefix="section")
    return title or cleaned, _finalize(spans)


def youtube_id(url: str) -> str | None:
    parsed = urlparse(url.strip())
    host = (parsed.hostname or "").lower()
    if host in {"youtu.be", "www.youtu.be"}:
        candidate = parsed.path.strip("/").split("/")[0]
        return candidate or None
    if host in {"youtube.com", "www.youtube.com", "m.youtube.com", "music.youtube.com"}:
        if parsed.path == "/watch":
            return parse_qs(parsed.query).get("v", [None])[0]
        for prefix in ("/shorts/", "/embed/", "/live/"):
            if parsed.path.startswith(prefix):
                return parsed.path[len(prefix) :].split("/")[0] or None
    return None


def _finalize(spans: list[dict]) -> list[dict]:
    ready = []
    for span in spans:
        text = (span.get("text") or "").replace("\x00", "").strip()
        if len(text) < 40:
            continue
        ready.append({"locator": span["locator"], "text": text[:8000]})
    if not ready:
        raise ParseError("No readable text was found. Scans and empty slides need a text layer or a transcript.")
    clipped = []
    total = 0
    for span in ready:
        if total >= MAX_TEXT_CHARS:
            break
        clipped.append(span)
        total += len(span["text"])
    return clipped


def _plain(text: str, locator_prefix: str) -> list[dict]:
    blocks = [block.strip() for block in re.split(r"\n\s*\n", text) if block.strip()]
    if not blocks:
        blocks = [text.strip()]
    spans = []
    bucket: list[str] = []
    size = 0
    index = 1
    for block in blocks:
        if bucket and size + len(block) > 1200:
            spans.append({"locator": f"{locator_prefix} {index}", "text": "\n\n".join(bucket)})
            index += 1
            bucket = []
            size = 0
        bucket.append(block)
        size += len(block)
    if bucket:
        spans.append({"locator": f"{locator_prefix} {index}", "text": "\n\n".join(bucket)})
    return spans


def _pdf(data: bytes) -> list[dict]:
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(data))
    spans = []
    for index, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""
        if text.strip():
            spans.append({"locator": f"p. {index}", "text": text})
    return spans


def _docx(data: bytes) -> list[dict]:
    from docx import Document

    document = Document(io.BytesIO(data))
    parts = [p.text.strip() for p in document.paragraphs if p.text and p.text.strip()]
    for table in document.tables:
        for row in table.rows:
            cells = [cell.text.strip() for cell in row.cells if cell.text and cell.text.strip()]
            if cells:
                parts.append(" | ".join(cells))
    return _plain("\n\n".join(parts), locator_prefix="section")


def _pptx(data: bytes) -> list[dict]:
    from pptx import Presentation

    presentation = Presentation(io.BytesIO(data))
    spans = []
    for index, slide in enumerate(presentation.slides, start=1):
        parts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = shape.text_frame.text.strip()
                if text:
                    parts.append(text)
        if parts:
            spans.append({"locator": f"slide {index}", "text": "\n".join(parts)})
    return spans


def _xlsx(data: bytes) -> list[dict]:
    from openpyxl import load_workbook

    workbook = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    spans = []
    try:
        for sheet in workbook.worksheets:
            rows = []
            for row in sheet.iter_rows(values_only=True):
                cells = [str(cell).strip() for cell in row if cell not in (None, "")]
                if cells:
                    rows.append(" | ".join(cells))
            if rows:
                spans.extend(_plain("\n".join(rows), locator_prefix=sheet.title or "sheet"))
    finally:
        workbook.close()
    return spans


def _html(html: str, locator_prefix: str) -> list[dict]:
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    for tag in soup(["script", "style", "noscript", "template"]):
        tag.decompose()
    text = soup.get_text("\n")
    return _plain(text, locator_prefix=locator_prefix)


def _reject_private_url(url: str) -> None:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.hostname:
        raise ParseError("Paste an http or https link.")
    _reject_host(parsed.hostname)


def _reject_host(hostname: str) -> None:
    host = hostname.strip().lower().rstrip(".")
    if host in {"localhost", "localhost.localdomain"} or host.endswith(".local"):
        raise ParseError("Private and local links are not fetched.")
    try:
        infos = socket.getaddrinfo(host, None)
    except socket.gaierror as exc:
        raise ParseError("That link could not be resolved.") from exc
    for info in infos:
        ip = ipaddress.ip_address(info[4][0])
        if (
            ip.is_private
            or ip.is_loopback
            or ip.is_link_local
            or ip.is_reserved
            or ip.is_multicast
            or ip.is_unspecified
        ):
            raise ParseError("Private and local links are not fetched.")


def _fetch_html(url: str) -> tuple[str, str]:
    current = url
    with httpx.Client(timeout=20.0, follow_redirects=False, headers={"User-Agent": "LynoteLearningPathOpen/1.0"}) as client:
        for _ in range(3):
            _reject_private_url(current)
            response = client.get(current)
            if response.status_code in {301, 302, 303, 307, 308}:
                location = response.headers.get("location")
                if not location:
                    break
                current = str(response.url.join(location))
                continue
            if response.status_code >= 400:
                raise ParseError(f"The link returned HTTP {response.status_code}.")
            content_type = response.headers.get("content-type", "")
            if "text" not in content_type and "html" not in content_type and "json" not in content_type:
                raise ParseError("That link is not a readable page. Upload a transcript or a document instead.")
            body = response.content[:2_000_000]
            title = ""
            if "html" in content_type:
                from bs4 import BeautifulSoup

                soup = BeautifulSoup(body, "html.parser")
                if soup.title and soup.title.string:
                    title = soup.title.string.strip()
            return title, body.decode(response.encoding or "utf-8", errors="replace")
    raise ParseError("The link redirected too many times.")


def _youtube_text(video_id: str) -> str:
    if not re.fullmatch(r"[A-Za-z0-9_-]{6,}", video_id):
        raise ParseError("That YouTube link is missing a video id.")
    try:
        from youtube_transcript_api import YouTubeTranscriptApi
    except ImportError as exc:
        raise ParseError("YouTube captions need the youtube-transcript-api package.") from exc
    try:
        api = YouTubeTranscriptApi()
        if hasattr(api, "fetch"):
            fetched = api.fetch(video_id)
            snippets: Iterable = getattr(fetched, "snippets", fetched)
            lines = []
            for snippet in snippets:
                if hasattr(snippet, "text"):
                    start = float(getattr(snippet, "start", 0) or 0)
                    text = snippet.text
                else:
                    start = float(snippet.get("start", 0) or 0)
                    text = snippet.get("text", "")
                if text and text.strip():
                    lines.append(f"[{_clock(start)}] {text.strip()}")
        else:
            rows = YouTubeTranscriptApi.get_transcript(video_id)
            lines = [f"[{_clock(float(row.get('start', 0)))}] {row.get('text', '').strip()}" for row in rows]
    except Exception as exc:  # library raises several public and private types
        raise ParseError(
            "No public captions were found for that video. Upload a .srt or .vtt transcript, "
            "or open the same link on Lynote, which can transcribe the lecture."
        ) from exc
    if not lines:
        raise ParseError("That video has no caption text.")
    return "\n".join(lines)


def _clock(seconds: float) -> str:
    whole = max(0, int(seconds))
    return f"{whole // 60:02d}:{whole % 60:02d}"
