from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass(slots=True, frozen=True)
class PdfWord:
    x0: float
    y0: float
    x1: float
    y1: float
    text: str
    block_number: int
    line_number: int
    word_number: int

    @property
    def vertical_center(self) -> float:
        return (self.y0 + self.y1) / 2.0


@dataclass(slots=True)
class PdfPageText:
    page_number: int
    text: str
    words: list[PdfWord] = field(default_factory=list)


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
            words = [
                PdfWord(
                    x0=float(item[0]),
                    y0=float(item[1]),
                    x1=float(item[2]),
                    y1=float(item[3]),
                    text=str(item[4]),
                    block_number=int(item[5]),
                    line_number=int(item[6]),
                    word_number=int(item[7]),
                )
                for item in page.get_text("words")
                if len(item) >= 8 and str(item[4]).strip()
            ]
            pages.append(
                PdfPageText(
                    page_number=index,
                    text=page.get_text("text"),
                    words=words,
                )
            )
    return pages
