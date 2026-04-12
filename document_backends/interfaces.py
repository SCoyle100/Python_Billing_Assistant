from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Protocol


class PdfToDocxConverter(Protocol):
    def convert_pdf_to_docx(self, input_path: str) -> str | None:
        ...


class MatrixPageMapper(Protocol):
    def read_page_markets(
        self,
        file_path: str,
        source_pdf_path: str | None = None,
    ) -> dict[int, Any]:
        ...


class MatrixDocumentRewriter(Protocol):
    def rewrite(self, file_path: str) -> None:
        ...


class InvoiceDataFrameBuilder(Protocol):
    def build(self, file_path: str):
        ...


class CapitolTableRebuilder(Protocol):
    def rebuild(self, file_path: str, invoices_list: list[tuple[Any, ...]]) -> None:
        ...


class DocxToPdfRenderer(Protocol):
    def render(self, docx_path: str) -> str | None:
        ...


class PdfToImageGenerator(Protocol):
    def generate(
        self,
        pdf_path: str,
        dpi: int = 600,
        vendor_name: str | None = None,
        invoice_data: list[tuple[Any, ...]] | None = None,
        page_market_mapping: dict[int, Any] | None = None,
    ) -> list[str]:
        ...


class DocxToImageGenerator(Protocol):
    def generate(
        self,
        docx_path: str,
        vendor_name: str,
        invoice_data: list[tuple[Any, ...]] | None = None,
        page_market_mapping: dict[int, Any] | None = None,
    ) -> list[str]:
        ...


@dataclass(frozen=True)
class MatrixVendorBackend:
    page_mapper: MatrixPageMapper
    document_rewriter: MatrixDocumentRewriter
    dataframe_builder: InvoiceDataFrameBuilder


@dataclass(frozen=True)
class CapitolVendorBackend:
    dataframe_builder: InvoiceDataFrameBuilder
    table_rebuilder: CapitolTableRebuilder


@dataclass(frozen=True)
class RenderingBackend:
    docx_to_pdf: DocxToPdfRenderer
    pdf_to_images: PdfToImageGenerator
    docx_to_images: DocxToImageGenerator


@dataclass(frozen=True)
class DocumentProcessingServices:
    pdf_to_docx: PdfToDocxConverter
    matrix: MatrixVendorBackend
    capitol: CapitolVendorBackend
    rendering: RenderingBackend
