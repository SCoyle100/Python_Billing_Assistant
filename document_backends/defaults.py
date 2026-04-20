from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from typing import Any

from image_generation.create_pdf_image import (
    convert_pdf_to_images,
    create_images_from_docx,
    create_pdf_from_docx,
)
from pdf_to_docx_ import PDFConverter
from vendor_invoice_logic.capitol_media_dataframe_1 import build_dataframe_from_capitol_media
from vendor_invoice_logic.capitol_media_rebuild import rebuild_capitol_media_table
from vendor_invoice_logic.matrix_media_dataframe import build_dataframe_from_word_document
from vendor_invoice_logic.matrix_media_pdf_market_map import read_page_markets_from_pdf
from .dotnet_document_tool import (
    convert_docx_to_pdf_with_syncfusion,
    rewrite_matrix_amounts_with_dotnet,
)

from .interfaces import (
    CapitolTableRebuilder,
    CapitolVendorBackend,
    DocxToImageGenerator,
    DocxToPdfRenderer,
    DocumentProcessingServices,
    InvoiceDataFrameBuilder,
    MatrixDocumentRewriter,
    MatrixPageMapper,
    MatrixVendorBackend,
    PdfToDocxConverter,
    PdfToImageGenerator,
    RenderingBackend,
)

logger = logging.getLogger(__name__)


@dataclass
class AdobePdfToDocxConverter(PdfToDocxConverter):
    converter: PDFConverter

    def convert_pdf_to_docx(self, input_path: str) -> str | None:
        return self.converter.convert_pdf_to_docx(input_path)


class WordComMatrixPageMapper(MatrixPageMapper):
    def read_page_markets(
        self,
        file_path: str,
        source_pdf_path: str | None = None,
    ) -> dict[int, Any]:
        try:
            from vendor_invoice_logic.matrix_media_market_map import read_page_markets
        except ImportError as exc:
            logger.warning("Word COM Matrix page mapper is unavailable: %s", exc)
            return {}

        return read_page_markets(file_path)


class PdfMatrixPageMapper(MatrixPageMapper):
    def read_page_markets(
        self,
        file_path: str,
        source_pdf_path: str | None = None,
    ) -> dict[int, Any]:
        if not source_pdf_path:
            return {}
        return read_page_markets_from_pdf(source_pdf_path)


@dataclass
class FallbackMatrixPageMapper(MatrixPageMapper):
    primary: MatrixPageMapper
    fallback: MatrixPageMapper

    def read_page_markets(
        self,
        file_path: str,
        source_pdf_path: str | None = None,
    ) -> dict[int, Any]:
        try:
            primary_result = self.primary.read_page_markets(
                file_path,
                source_pdf_path=source_pdf_path,
            )
            if primary_result:
                logger.info(
                    "Matrix page mapper used primary backend %s for %s",
                    type(self.primary).__name__,
                    source_pdf_path or file_path,
                )
                return primary_result
            logger.info(
                "Matrix page mapper primary backend %s returned no mapping for %s; falling back to %s",
                type(self.primary).__name__,
                source_pdf_path or file_path,
                type(self.fallback).__name__,
            )
        except Exception as exc:
            logger.warning(
                "Matrix page mapper primary backend %s failed for %s: %s. Falling back to %s",
                type(self.primary).__name__,
                source_pdf_path or file_path,
                exc,
                type(self.fallback).__name__,
            )

        return self.fallback.read_page_markets(
            file_path,
            source_pdf_path=source_pdf_path,
        )


class WordComMatrixDocumentRewriter(MatrixDocumentRewriter):
    def rewrite(self, file_path: str, page_market_mapping: dict[int, Any] | None = None) -> None:
        from vendor_invoice_logic.matrix_media_logic import analyze_word_document

        analyze_word_document(file_path, page_market_mapping=page_market_mapping)


class DotNetMatrixDocumentRewriter(MatrixDocumentRewriter):
    def rewrite(self, file_path: str, page_market_mapping: dict[int, Any] | None = None) -> None:
        if not rewrite_matrix_amounts_with_dotnet(file_path, page_market_mapping=page_market_mapping):
            raise RuntimeError(f".NET Matrix document rewriter failed for {file_path}")


@dataclass
class FallbackMatrixDocumentRewriter(MatrixDocumentRewriter):
    primary: MatrixDocumentRewriter
    fallback: MatrixDocumentRewriter

    def rewrite(self, file_path: str, page_market_mapping: dict[int, Any] | None = None) -> None:
        if os.getenv("MATRIX_REWRITER", "").lower() == "word":
            self.fallback.rewrite(file_path, page_market_mapping=page_market_mapping)
            return

        try:
            self.primary.rewrite(file_path, page_market_mapping=page_market_mapping)
            logger.info(
                "Matrix document rewriter used primary backend %s for %s",
                type(self.primary).__name__,
                file_path,
            )
            return
        except Exception as exc:
            logger.warning(
                "Matrix document rewriter primary backend %s failed for %s: %s. Falling back to %s",
                type(self.primary).__name__,
                file_path,
                exc,
                type(self.fallback).__name__,
            )

        self.fallback.rewrite(file_path, page_market_mapping=page_market_mapping)


@dataclass
class FunctionDataFrameBuilder(InvoiceDataFrameBuilder):
    builder: Any

    def build(self, file_path: str):
        return self.builder(file_path)


class FunctionCapitolTableRebuilder(CapitolTableRebuilder):
    def rebuild(self, file_path: str, invoices_list: list[tuple[Any, ...]]) -> None:
        rebuild_capitol_media_table(file_path, invoices_list)


class WordComDocxToPdfRenderer(DocxToPdfRenderer):
    def render(self, docx_path: str) -> str | None:
        return create_pdf_from_docx(docx_path)


class SyncfusionDocxToPdfRenderer(DocxToPdfRenderer):
    def render(self, docx_path: str) -> str | None:
        return convert_docx_to_pdf_with_syncfusion(docx_path)


class FitzPdfToImageGenerator(PdfToImageGenerator):
    def generate(
        self,
        pdf_path: str,
        dpi: int = 600,
        vendor_name: str | None = None,
        invoice_data: list[tuple[Any, ...]] | None = None,
        page_market_mapping: dict[int, Any] | None = None,
    ) -> list[str]:
        return convert_pdf_to_images(
            pdf_path,
            dpi=dpi,
            vendor_name=vendor_name,
            invoice_data=invoice_data,
            page_market_mapping=page_market_mapping,
        )


class WordComDocxToImageGenerator(DocxToImageGenerator):
    def generate(
        self,
        docx_path: str,
        vendor_name: str,
        invoice_data: list[tuple[Any, ...]] | None = None,
        page_market_mapping: dict[int, Any] | None = None,
    ) -> list[str]:
        return create_images_from_docx(
            docx_path,
            vendor_name,
            invoice_data=invoice_data,
            page_market_mapping=page_market_mapping,
        )


@dataclass
class RenderedPdfDocxToImageGenerator(DocxToImageGenerator):
    docx_to_pdf: DocxToPdfRenderer
    pdf_to_images: PdfToImageGenerator

    def generate(
        self,
        docx_path: str,
        vendor_name: str,
        invoice_data: list[tuple[Any, ...]] | None = None,
        page_market_mapping: dict[int, Any] | None = None,
    ) -> list[str]:
        pdf_path = self.docx_to_pdf.render(docx_path)
        if not pdf_path:
            return []

        return self.pdf_to_images.generate(
            pdf_path,
            dpi=600,
            vendor_name=vendor_name,
            invoice_data=invoice_data,
            page_market_mapping=page_market_mapping,
        )


def build_default_document_services() -> DocumentProcessingServices:
    pdf_to_images = FitzPdfToImageGenerator()
    if os.getenv("SYNCFUSION_LICENSE_KEY"):
        docx_to_pdf = SyncfusionDocxToPdfRenderer()
        docx_to_images = RenderedPdfDocxToImageGenerator(
            docx_to_pdf=docx_to_pdf,
            pdf_to_images=pdf_to_images,
        )
    else:
        docx_to_pdf = WordComDocxToPdfRenderer()
        docx_to_images = WordComDocxToImageGenerator()

    return DocumentProcessingServices(
        pdf_to_docx=AdobePdfToDocxConverter(converter=PDFConverter()),
        matrix=MatrixVendorBackend(
            page_mapper=FallbackMatrixPageMapper(
                primary=PdfMatrixPageMapper(),
                fallback=WordComMatrixPageMapper(),
            ),
            document_rewriter=FallbackMatrixDocumentRewriter(
                primary=DotNetMatrixDocumentRewriter(),
                fallback=WordComMatrixDocumentRewriter(),
            ),
            dataframe_builder=FunctionDataFrameBuilder(builder=build_dataframe_from_word_document),
        ),
        capitol=CapitolVendorBackend(
            dataframe_builder=FunctionDataFrameBuilder(builder=build_dataframe_from_capitol_media),
            table_rebuilder=FunctionCapitolTableRebuilder(),
        ),
        rendering=RenderingBackend(
            docx_to_pdf=docx_to_pdf,
            pdf_to_images=pdf_to_images,
            docx_to_images=docx_to_images,
        ),
    )
