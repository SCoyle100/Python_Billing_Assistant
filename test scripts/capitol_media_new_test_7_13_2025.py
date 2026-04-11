"""
combined_capitol_to_invoice_single.py
-------------------------------------------------
• Prompts for ONE Word document that already contains the
  city/amount lines and the table you want to overwrite.
• Extracts the rows, builds a DataFrame, then rewrites
  that same table (saving a *_updated.docx* copy).

Dependencies:
  python-docx, pandas, win32com.client, openai (with OPENAI_API_KEY)
  and Word installed on Windows.
"""
import os, re, logging
from copy import deepcopy
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd

from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from vendor_invoice_logic.capitol_media_dataframe_1 import build_dataframe_from_capitol_media


# ─────────────── logging setup ───────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s | %(message)s"
)
log = logging.getLogger(__name__)


# ─────────────── utilities ───────────────
def parse_dollar_amount(s: str) -> float:
    return float(re.sub(r'[^\d\.\-]', '', s)) if re.search(r'\d', s) else 0.0

def format_amount(x: float) -> str:
    sign = '-' if x < 0 else ''
    return f"{sign}{abs(x):,.2f}"


# ─────────────── extraction ───────────────
def _pick_invoice_table(doc):
    """
    Return the first python-docx table whose first *visible* row contains
    both 'Description' and 'Amount' (case-insensitive).  Falls back to the
    last table in the document.
    """
    for t in doc.tables:
        header = " ".join(c.text.lower() for c in t.rows[0].cells)
        if "description" in header and "amount" in header:
            return t
    return doc.tables[-1] if doc.tables else None


def build_dataframe(path: str) -> pd.DataFrame:
    """
    Use the shared Capitol Media extractor and reshape it for this table updater.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(path)

    df = build_dataframe_from_capitol_media(path).rename(
        columns={"Market": "Description"}
    )
    if "Amount" in df.columns:
        df["Amount"] = df["Amount"].apply(parse_dollar_amount)
    else:
        df["Amount"] = pd.Series(dtype=float)

    log.info("Extracted %d row(s):\n%s",
             len(df), df.to_string(index=False) if not df.empty else "<empty>")
    return df



# ─────────────── table update ───────────────
def update_table(doc_path: str, df: pd.DataFrame) -> str:
    doc = Document(doc_path)
    tbl = _pick_invoice_table(doc)          # reuse the same picker
    if tbl is None:
        raise RuntimeError("No table with 'Description' / 'Amount' header found.")

    xml_tbl   = tbl._tbl
    kids      = list(xml_tbl)               # [tblPr, tblGrid, <tr> …]
    xml_start = 2                           # first <w:tr>
    template  = kids[xml_start + 2]         # clone row-2 as pattern
    data_idx  = xml_start + 3               # where the city rows begin

    # wipe all old data rows
    for tr in kids[data_idx:]:
        xml_tbl.remove(tr)

    # rebuild from DataFrame
    for n, row in enumerate(df.itertuples(index=False)):
        tr = deepcopy(template)
        cells = tr.findall(qn("w:tc"))

        cells[0].xpath(".//w:t")[0].text = str(row.Description)

        num_txt = f"{abs(float(row.Amount)):,.2f}"
        if float(row.Amount) < 0:
            num_txt = "-" + num_txt
        cells[1].xpath(".//w:t")[0].text = num_txt

        xml_tbl.insert(data_idx + n, tr)

    out = os.path.splitext(doc_path)[0] + "_updated.docx"
    doc.save(out)
    return out



# ─────────────── main flow ───────────────
def main():
    root = tk.Tk(); root.withdraw()

    doc_path = filedialog.askopenfilename(
        title="Select the Capitol Media invoice (.docx)",
        filetypes=[("Word files", "*.docx")]
    )
    if not doc_path:
        log.info("No file selected; exiting.")
        return

    try:
        df = build_dataframe(doc_path)
        if df.empty:
            messagebox.showerror("No data", "Could not find city/amount rows.")
            return
    except Exception as e:
        log.exception("Error during extraction")
        messagebox.showerror("Extraction error", str(e))
        return

    try:
        new_doc = update_table(doc_path, df)
        messagebox.showinfo("Success", f"Invoice updated:\n{new_doc}")
    except Exception as e:
        log.exception("Error during table update")
        messagebox.showerror("Update error", str(e))


if __name__ == "__main__":
    main()
