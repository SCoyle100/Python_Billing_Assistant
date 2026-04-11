from __future__ import annotations

from copy import deepcopy
import logging
import re

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


HEADER_KEYWORDS = {"description", "amount"}


def _parse_amount(value):
    cleaned = re.sub(r"[^\d.\-]", "", str(value or ""))
    if not cleaned:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def _format_currency(amount):
    return f"${amount:,.2f}"


def _split_invoice_rows(invoice_rows):
    """
    Convert raw Capitol invoice rows into deterministic display rows.
    Large amounts are split into PART rows capped at $5,000 each.
    """
    display_rows = []
    running_total = 0.0

    for market, amount in invoice_rows:
        normalized_market = str(market or "").strip()
        normalized_amount = _parse_amount(amount)

        if not normalized_market or normalized_amount <= 0:
            continue

        remaining = normalized_amount
        part_index = 0
        while remaining > 0:
            part_index += 1
            part_amount = min(5000.0, remaining)
            remaining -= part_amount

            if normalized_amount > 5000.0:
                description = f"{normalized_market} - PART {chr(64 + part_index)}"
            else:
                description = normalized_market

            display_rows.append((description, _format_currency(part_amount)))
            running_total += part_amount

    return display_rows, running_total


def _is_discount_line(text):
    lowered = str(text or "").strip().lower()
    return "discount" in lowered or "markup" in lowered


def _looks_like_intro_line(text):
    lowered = str(text or "").strip().lower()
    if not lowered or _is_discount_line(lowered):
        return False

    if any(token in lowered for token in ["invoice", "start", "weeks", "#", "/"]):
        return True

    if lowered.count(",") >= 3:
        return True

    return False


def _row_texts(row):
    texts = []
    for cell in row.cells:
        cell_text = " ".join(
            paragraph.text.strip()
            for paragraph in cell.paragraphs
            if paragraph.text and paragraph.text.strip()
        )
        texts.append(cell_text)
    return texts


def _find_main_table(doc):
    for table in doc.tables:
        for row in table.rows:
            lowered = {text.strip().lower() for text in _row_texts(row) if text.strip()}
            if HEADER_KEYWORDS.issubset(lowered):
                return table

    best_table = None
    best_score = -1
    for table in doc.tables:
        score = sum(bool(re.search(r"\d", text)) for row in table.rows for text in _row_texts(row))
        if score > best_score:
            best_table = table
            best_score = score
    return best_table


def _find_header_row_index(table):
    for index, row in enumerate(table.rows):
        lowered = {text.strip().lower() for text in _row_texts(row) if text.strip()}
        if HEADER_KEYWORDS.issubset(lowered):
            return index
    return None


def _find_total_row_index(table, start_index):
    for index in range(start_index, len(table.rows)):
        row_text = " ".join(_row_texts(table.rows[index])).lower()
        if "total" in row_text:
            return index
    return len(table.rows) - 1


def _first_non_empty_paragraph_text(cell):
    for paragraph in cell.paragraphs:
        text = paragraph.text.strip()
        if text:
            return text
    return ""


def _paragraph_lines(cell):
    return [
        paragraph.text.strip()
        for paragraph in cell.paragraphs
        if paragraph.text and paragraph.text.strip()
    ]


def _extract_intro_lines(detail_row):
    intro_lines = []
    for line in _paragraph_lines(detail_row.cells[0]):
        if _looks_like_intro_line(line):
            intro_lines.append(line)
            continue
        break
    return intro_lines


def _clear_cell(cell):
    tc = cell._tc
    for child in list(tc):
        if child.tag.endswith("tcPr"):
            continue
        tc.remove(child)


def _add_lines_to_cell(
    cell,
    lines,
    *,
    font_name="Arial",
    font_size=9,
    alignment=WD_ALIGN_PARAGRAPH.LEFT,
    bold=False,
):
    _clear_cell(cell)

    normalized_lines = list(lines) if lines else [""]
    for line_index, line in enumerate(normalized_lines):
        paragraph = cell.add_paragraph() if line_index > 0 else cell.add_paragraph()
        paragraph.alignment = alignment
        run = paragraph.add_run(str(line))
        run.font.name = font_name
        run.font.size = Pt(font_size)
        run.bold = bold


def _populate_invoice_row(row, description_lines, amount_text=""):
    _add_lines_to_cell(
        row.cells[0],
        description_lines,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
    )
    _add_lines_to_cell(
        row.cells[-1],
        [amount_text] if amount_text else [""],
        alignment=WD_ALIGN_PARAGRAPH.RIGHT,
    )


def _remove_row_height(row):
    tr = row._tr
    tr_pr = tr.trPr
    if tr_pr is None:
        return

    for child in list(tr_pr):
        if child.tag == qn("w:trHeight"):
            tr_pr.remove(child)


def _set_row_min_height(row, value=360):
    tr = row._tr
    tr_pr = tr.get_or_add_trPr()
    tr_height = OxmlElement("w:trHeight")
    tr_height.set(qn("w:val"), str(value))
    tr_height.set(qn("w:hRule"), "atLeast")
    tr_pr.append(tr_height)


def rebuild_capitol_media_table(docx_path, invoice_rows):
    """
    Preserve the converted Word doc's stacked header area, but regenerate the
    detail and total rows under the Description/Amount header.
    """
    doc = Document(docx_path)
    table = _find_main_table(doc)
    if table is None:
        raise ValueError(f"No Capitol Media invoice table found in {docx_path}")

    header_row_index = _find_header_row_index(table)
    if header_row_index is None:
        raise ValueError(f"Could not locate Description/Amount header row in {docx_path}")

    detail_row_index = header_row_index + 1
    if detail_row_index >= len(table.rows):
        raise ValueError(f"Could not locate detail row after Description/Amount header in {docx_path}")

    total_row_index = _find_total_row_index(table, detail_row_index)

    original_detail_row = table.rows[detail_row_index]
    original_total_row = table.rows[total_row_index]
    preserved_intro_lines = _extract_intro_lines(original_detail_row)

    detail_template = deepcopy(original_detail_row._tr)
    total_template = deepcopy(original_total_row._tr)

    for row_index in range(total_row_index, detail_row_index - 1, -1):
        table._tbl.remove(table.rows[row_index]._tr)

    detail_rows, running_total = _split_invoice_rows(invoice_rows)
    row_specs = []

    if preserved_intro_lines:
        row_specs.append((preserved_intro_lines, ""))

    for description, amount in detail_rows:
        row_specs.append(([description], amount))

    row_templates = [deepcopy(detail_template) for _ in row_specs]
    row_templates.append(deepcopy(total_template))

    header_tr = table.rows[header_row_index]._tr
    for template in reversed(row_templates):
        header_tr.addnext(template)

    inserted_rows_start = detail_row_index
    for offset, (description_lines, amount_text) in enumerate(row_specs):
        row = table.rows[inserted_rows_start + offset]
        _remove_row_height(row)
        _set_row_min_height(row, 360 if amount_text else 520)
        _populate_invoice_row(row, description_lines, amount_text)

    total_row = table.rows[inserted_rows_start + len(row_specs)]
    _remove_row_height(total_row)
    _set_row_min_height(total_row, 420)
    _add_lines_to_cell(
        total_row.cells[0],
        ["Total"],
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
        bold=True,
    )
    _add_lines_to_cell(
        total_row.cells[-1],
        [_format_currency(running_total)],
        alignment=WD_ALIGN_PARAGRAPH.RIGHT,
        bold=True,
    )

    doc.save(docx_path)
    logging.info("Rebuilt Capitol Media pricing table in %s", docx_path)
