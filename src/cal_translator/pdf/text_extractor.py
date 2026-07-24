from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class PdfPageText:
    page_number: int
    text: str


def extract_pdf_text(pdf_path: Path) -> list[PdfPageText]:
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF file not found: {pdf_path}")

    try:
        import fitz
    except ImportError as exc:
        raise RuntimeError(
            "PyMuPDF is required. Install the project with: python -m pip install -e ."
        ) from exc

    pages: list[PdfPageText] = []
    with fitz.open(pdf_path) as document:
        for index, page in enumerate(document, start=1):
            pages.append(PdfPageText(page_number=index, text=page.get_text("text")))
    return pages
