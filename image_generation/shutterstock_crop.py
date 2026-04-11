"""
invoice_cropper.py
------------------
Pick a PDF or image, extract the invoice area, and save the result to an
/images folder next to this script.

Requires:
    pip install pymupdf opencv-python pillow numpy
"""

import fitz                     # PyMuPDF
import cv2
import numpy as np
from PIL import Image
from tkinter import Tk, filedialog, messagebox
from pathlib import Path
import os


# ──────────────────────────────────────────────────────────────────────────────
def crop_file() -> None:
    root = Tk(); root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select a File (PDF or Image)",
        filetypes=[("PDF & Image Files",
                    "*.pdf;*.png;*.jpg;*.jpeg;*.tif;*.tiff"),
                   ("All Files", "*.*")]
    )
    if not file_path:
        messagebox.showerror("Error", "No file selected.")
        return
    if not os.path.exists(file_path):
        messagebox.showerror("Error", f"File not found: {file_path}")
        return

    try:
        ext = Path(file_path).suffix.lower()
        if ext == ".pdf":
            process_pdf(file_path)
        elif ext in {".png", ".jpg", ".jpeg", ".tif", ".tiff"}:
            process_image(file_path)
        else:
            messagebox.showerror("Error", "Unsupported file type.")
            return

        messagebox.showinfo("Done", "Cropped image(s) saved in the /images folder.")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")


# ──────────────────────────────────────────────────────────────────────────────
def process_pdf(file_path: str) -> None:
    pdf        = fitz.open(file_path)
    script_dir = Path(__file__).resolve().parent
    out_dir    = script_dir / "images"
    out_dir.mkdir(parents=True, exist_ok=True)

    mat = fitz.Matrix(300 / 72, 300 / 72)            # ≈300 dpi
    for i, page in enumerate(pdf, start=1):
        pix = page.get_pixmap(matrix=mat, alpha=False)
        w, h = pix.width, pix.height
        img  = Image.frombytes("RGB", (w, h), pix.samples)
        bgr  = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
        crop = crop_image(bgr)

        cv2.imwrite(str(out_dir / f"invoice_{i}.png"), crop)


def create_cropped_shutterstock_image(pdf_path: str, output_path: str, page_index: int = 0) -> str:
    """
    Render a PDF page, crop it using the Shutterstock invoice crop logic,
    upscale it for Word insertion, and save to the requested output path.
    """
    pdf = fitz.open(pdf_path)
    try:
        if pdf.page_count == 0:
            raise ValueError(f"PDF has no pages: {pdf_path}")
        if page_index < 0 or page_index >= pdf.page_count:
            raise IndexError(f"Requested page_index {page_index} is out of range for {pdf_path}")

        page = pdf.load_page(page_index)
        mat = fitz.Matrix(300 / 72, 300 / 72)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        image = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        bgr = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)
        crop = crop_image(bgr)
        sharp = upscale(crop, factor=6)

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        cv2.imwrite(str(output_file), sharp)
        return str(output_file)
    finally:
        pdf.close()


# ──────────────────────────────────────────────────────────────────────────────
def process_image(file_path: str) -> None:
    script_dir = Path(__file__).resolve().parent
    out_dir    = script_dir / "images"
    out_dir.mkdir(exist_ok=True)

    img   = cv2.imread(file_path)                    # BGR already
    crop  = crop_image(img)
    sharp = upscale(crop, factor=6)

    cv2.imwrite(str(out_dir / f"cropped_{Path(file_path).name}"), sharp)


# ──────────────────────────────────────────────────────────────────────────────
def crop_image(img: np.ndarray, *,
               bottom_cut_ratio: float = 0.15, pad: int = 20) -> np.ndarray:
    """
    Grab everything that isn't blank page, ignoring the bottom 15 % where the
    red Print / Close buttons live.
    """
    gray  = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray  = cv2.GaussianBlur(gray, (5, 5), 0)

    # Otsu gives robust text/background split
    _, thresh = cv2.threshold(gray, 0, 255,
                              cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

    h_total   = img.shape[0]
    top_area  = thresh[: int(h_total * (1 - bottom_cut_ratio)), :]

    coords = cv2.findNonZero(top_area)               # all non-black pixels
    if coords is None:
        return img                                   # nothing found → keep page

    x, y, w, h = cv2.boundingRect(coords)
    x0 = max(x - pad, 0)
    y0 = max(y - pad, 0)
    x1 = min(x + w + pad, img.shape[1])
    y1 = min(y + h + pad, img.shape[0])

    return img[y0:y1, x0:x1]


def upscale(img: np.ndarray, *, factor: int = 6) -> np.ndarray:
    """Enlarge for on-screen sharpness."""
    return cv2.resize(img, None, fx=factor, fy=factor,
                      interpolation=cv2.INTER_LANCZOS4)


# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    crop_file()
