import logging
import re
from typing import Callable


logging.basicConfig(level=logging.INFO)

SERVICE_PERIOD_RE = re.compile(
    r"\b\d{1,2}/\d{1,2}/\d{2,4}\s*[-–]\s*\d{1,2}/\d{1,2}/\d{2,4}\b"
)
MARKET_LINE_RE = re.compile(r"^[A-Za-z][A-Za-z0-9 .&()'/-]{1,80}$")


def _extract_text_with_fitz_per_page(pdf_path: str) -> list[str]:
    import fitz

    document = fitz.open(str(pdf_path))
    try:
        return [page.get_text("text") or "" for page in document]
    finally:
        document.close()


def _extract_text_with_pypdf2_per_page(pdf_path: str) -> list[str]:
    from PyPDF2 import PdfReader

    reader = PdfReader(str(pdf_path))
    return [(page.extract_text() or "") for page in reader.pages]


def extract_pdf_pages(pdf_path: str) -> list[str]:
    extractors: tuple[Callable[[str], list[str]], ...] = (
        _extract_text_with_fitz_per_page,
        _extract_text_with_pypdf2_per_page,
    )
    last_error = None

    for extractor in extractors:
        try:
            pages = extractor(pdf_path)
            if pages:
                return pages
        except Exception as exc:
            last_error = exc

    if last_error:
        raise last_error

    return []


def normalize_market_name(market: str) -> str:
    market = " ".join(str(market or "").split()).strip()
    lowered = market.lower().replace(".", "")
    if lowered in {"fort payne", "ft payne"}:
        return "Fort Payne"
    return market


def _clean_page_lines(page_text: str) -> list[str]:
    lines = []
    for raw_line in (page_text or "").splitlines():
        line = " ".join(raw_line.split()).strip()
        if line:
            lines.append(line)
    return lines


def _extract_service_period(lines: list[str]) -> str:
    for line in lines:
        match = SERVICE_PERIOD_RE.search(line)
        if match:
            return match.group(0)
    return ""


def _is_market_candidate(line: str) -> bool:
    lowered = line.lower()
    if lowered in {"market", "amount", "description", "service period", "qty"}:
        return False
    if "invoice" in lowered or "total" in lowered:
        return False
    if SERVICE_PERIOD_RE.search(line):
        return False
    if "$" in line:
        return False
    if not MARKET_LINE_RE.match(line):
        return False
    return any(char.isalpha() for char in line)


def _extract_market(lines: list[str]) -> str:
    market_header_seen = False

    for line in lines:
        lowered = line.lower()

        if lowered == "market":
            market_header_seen = True
            continue

        if market_header_seen and _is_market_candidate(line):
            return normalize_market_name(line)

    for line in lines:
        if _is_market_candidate(line):
            return normalize_market_name(line)

    return ""


def read_page_markets_from_pdf(pdf_path: str) -> dict[int, tuple[str, str]]:
    """
    Best-effort PDF-first page mapper for Matrix Media source PDFs.
    This is intentionally conservative and should be used with a fallback mapper.
    """
    page_texts = extract_pdf_pages(pdf_path)
    page_meta: dict[int, tuple[str, str]] = {}

    for page_index, page_text in enumerate(page_texts, start=1):
        lines = _clean_page_lines(page_text)
        if not lines:
            continue

        market = _extract_market(lines)
        service_period = _extract_service_period(lines)

        if market:
            page_meta[page_index] = (market, service_period)
            logging.info(
                "PDF page mapper: page %s -> market '%s', service period '%s'",
                page_index,
                market,
                service_period,
            )

    return page_meta
