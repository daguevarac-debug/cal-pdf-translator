from __future__ import annotations

import argparse
import logging
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import fitz  # PyMuPDF
from openpyxl import load_workbook


LOGGER = logging.getLogger(__name__)

REQUIRED_COLUMNS = [
    "unit_id",
    "page",
    "action",
    "original_text",
    "english_translation",
    "bbox_x0",
    "bbox_y0",
    "bbox_x1",
    "bbox_y1",
    "font",
    "font_size",
    "color",
    "flags",
    "source_span_ids",
    "block_number",
]

TRANSLATION_NOTE = "Unofficial English translation - the original Spanish certificate prevails."


@dataclass(slots=True)
class TranslationUnitRow:
    unit_id: int
    page: int
    action: str
    original_text: str
    english_translation: str
    bbox_x0: float
    bbox_y0: float
    bbox_x1: float
    bbox_y1: float
    font: str
    font_size: float
    color: int
    flags: int
    source_span_ids: str
    block_number: int

    @property
    def bbox(self) -> fitz.Rect:
        return fitz.Rect(self.bbox_x0, self.bbox_y0, self.bbox_x1, self.bbox_y1)


@dataclass(slots=True)
class Summary:
    translate_processed: int = 0
    remove_processed: int = 0
    keep_ignored: int = 0
    empty_translations: int = 0
    reduced_font_texts: int = 0
    not_fit_texts: int = 0
    row_failures: int = 0


@dataclass(slots=True)
class FontSpec:
    fontname: str
    fontfile: str | None


def setup_logging() -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Apply reviewed translations onto a source PDF.")
    parser.add_argument("pdf_path", type=Path, help="Path to original Spanish PDF")
    parser.add_argument("excel_path", type=Path, help="Path to reviewed Excel file")
    parser.add_argument("--sheet", default="translation_ready", help="Worksheet name with translation units")
    parser.add_argument("--output", type=Path, default=None, help="Output PDF path")
    return parser.parse_args()


def to_project_path(project_root: Path, path_value: Path) -> Path:
    if path_value.is_absolute():
        return path_value
    return (project_root / path_value).resolve()


def normalize_font_family(font_name: str) -> str:
    cleaned = font_name.lower().strip()
    if "+" in cleaned:
        cleaned = cleaned.split("+", maxsplit=1)[1]
    cleaned = cleaned.replace(" ", "").replace("_", "").replace("-", "")
    for suffix in ("psmt", "mt", "std", "regular"):
        if cleaned.endswith(suffix):
            cleaned = cleaned[: -len(suffix)]
    return cleaned


def is_bold(font_name: str, flags: int) -> bool:
    return "bold" in font_name.lower() or bool(flags & (1 << 4))


def is_italic(font_name: str, flags: int) -> bool:
    lower = font_name.lower()
    return "italic" in lower or "oblique" in lower or bool(flags & (1 << 1))


def choose_font_spec(unit: TranslationUnitRow, arial_paths: dict[str, Path]) -> FontSpec:
    source_font = unit.font
    bold = is_bold(source_font, unit.flags)
    italic = is_italic(source_font, unit.flags)
    normalized = normalize_font_family(source_font)

    if normalized == "helvetica":
        if bold and italic:
            return FontSpec("hebi", None)
        if bold:
            return FontSpec("hebo", None)
        if italic:
            return FontSpec("heit", None)
        return FontSpec("helv", None)

    if bold and italic:
        key = "arialbi"
        fallback = "hebi"
    elif bold:
        key = "arialbd"
        fallback = "hebo"
    elif italic:
        key = "ariali"
        fallback = "heit"
    else:
        key = "arial"
        fallback = "helv"

    arial_path = arial_paths.get(key)
    if arial_path is not None and arial_path.exists():
        return FontSpec(fontname=f"f_{key}", fontfile=str(arial_path))
    return FontSpec(fontname=fallback, fontfile=None)


def parse_int(value: Any, field_name: str) -> int:
    if value is None:
        raise ValueError(f"Missing value for {field_name}")
    if isinstance(value, bool):
        raise ValueError(f"Invalid boolean for {field_name}")
    return int(value)


def parse_float(value: Any, field_name: str) -> float:
    if value is None:
        raise ValueError(f"Missing value for {field_name}")
    return float(value)


def parse_row(row_values: dict[str, Any]) -> TranslationUnitRow:
    action = str(row_values["action"] or "").strip().upper()
    return TranslationUnitRow(
        unit_id=parse_int(row_values["unit_id"], "unit_id"),
        page=parse_int(row_values["page"], "page"),
        action=action,
        original_text=str(row_values["original_text"] or ""),
        english_translation=str(row_values["english_translation"] or ""),
        bbox_x0=parse_float(row_values["bbox_x0"], "bbox_x0"),
        bbox_y0=parse_float(row_values["bbox_y0"], "bbox_y0"),
        bbox_x1=parse_float(row_values["bbox_x1"], "bbox_x1"),
        bbox_y1=parse_float(row_values["bbox_y1"], "bbox_y1"),
        font=str(row_values["font"] or ""),
        font_size=parse_float(row_values["font_size"], "font_size"),
        color=parse_int(row_values["color"], "color"),
        flags=parse_int(row_values["flags"], "flags"),
        source_span_ids=str(row_values["source_span_ids"] or ""),
        block_number=parse_int(row_values["block_number"], "block_number"),
    )


def read_translation_units(excel_path: Path, sheet_name: str) -> list[TranslationUnitRow]:
    try:
        workbook = load_workbook(excel_path, data_only=True)
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"Could not open Excel file: {excel_path} | {exc}") from exc

    if sheet_name not in workbook.sheetnames:
        raise RuntimeError(f"Worksheet '{sheet_name}' not found in {excel_path}")

    worksheet = workbook[sheet_name]
    header_cells = next(worksheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
    if header_cells is None:
        raise RuntimeError(f"Worksheet '{sheet_name}' has no header row")

    header_map: dict[str, int] = {}
    for index, value in enumerate(header_cells):
        if value is None:
            continue
        header_map[str(value).strip()] = index

    missing_columns = [column for column in REQUIRED_COLUMNS if column not in header_map]
    if missing_columns:
        raise RuntimeError(f"Missing required columns in sheet '{sheet_name}': {missing_columns}")

    rows: list[TranslationUnitRow] = []
    for excel_row in worksheet.iter_rows(min_row=2, values_only=True):
        if all(cell is None for cell in excel_row):
            continue

        row_values = {column: excel_row[header_map[column]] for column in REQUIRED_COLUMNS}
        rows.append(parse_row(row_values))

    return rows


def is_table_header_candidate(unit: TranslationUnitRow) -> bool:
    if not is_bold(unit.font, unit.flags):
        return False
    if unit.page not in (2, 3):
        return False
    text = unit.original_text.strip()
    if text == "":
        return False
    if len(text) > 40:
        return False
    if re.match(r"^\s*\d", text):
        return False
    return True


def is_centered_unit(unit: TranslationUnitRow, page_rect: fitz.Rect) -> bool:
    normalized = unit.original_text.strip().lower()
    if normalized in {"certificado de calibraciÃ³n", "fin del certificado"}:
        return True
    if is_table_header_candidate(unit):
        return True
    unit_center_x = (unit.bbox_x0 + unit.bbox_x1) / 2.0
    page_center_x = (page_rect.x0 + page_rect.x1) / 2.0
    return abs(unit_center_x - page_center_x) < 25.0


def should_expand_text_box(unit: TranslationUnitRow) -> bool:
    text = unit.original_text.strip()
    return is_table_header_candidate(unit) and len(text) <= 30


def build_text_rect(unit: TranslationUnitRow) -> fitz.Rect:
    rect = fitz.Rect(unit.bbox)
    if should_expand_text_box(unit):
        rect = fitz.Rect(rect.x0 - 2.0, rect.y0, rect.x1 + 2.0, rect.y1)
    return rect


def min_font_size_for(unit: TranslationUnitRow) -> float:
    if unit.font_size < 5.0:
        return 3.2
    return 4.0


def alignment_for(unit: TranslationUnitRow, page_rect: fitz.Rect) -> int:
    text = unit.english_translation
    if len(text) > 100:
        return fitz.TEXT_ALIGN_JUSTIFY
    if is_centered_unit(unit, page_rect):
        return fitz.TEXT_ALIGN_CENTER
    return fitz.TEXT_ALIGN_LEFT


def try_insert_translation(
    page: fitz.Page,
    unit: TranslationUnitRow,
    font_spec: FontSpec,
    summary: Summary,
) -> None:
    text_rect = build_text_rect(unit)
    original_size = max(0.1, unit.font_size)
    current_size = original_size
    min_size = min_font_size_for(unit)
    reduction_step = 0.25
    page_rect = page.rect
    align = alignment_for(unit, page_rect)
    has_more_than_100 = len(unit.english_translation) > 100
    reduced = False

    while current_size >= min_size:
        result = page.insert_textbox(
            text_rect,
            unit.english_translation,
            fontsize=current_size,
            fontname=font_spec.fontname,
            fontfile=font_spec.fontfile,
            color=fitz.sRGB_to_pdf(unit.color),
            align=align,
        )
        if result >= 0:
            if reduced:
                summary.reduced_font_texts += 1
            return

        reduced = True
        if has_more_than_100 and current_size > original_size * 0.9:
            current_size = max(original_size * 0.9, current_size - reduction_step)
        else:
            current_size -= reduction_step

    summary.not_fit_texts += 1
    LOGGER.warning(
        "Text does not fit | unit_id=%s | page=%s | text=%s",
        unit.unit_id,
        unit.page,
        unit.english_translation,
    )


def add_redaction(page: fitz.Page, unit: TranslationUnitRow) -> None:
    rect = fitz.Rect(unit.bbox)
    redaction_rect = fitz.Rect(rect.x0 - 0.4, rect.y0 - 0.4, rect.x1 + 0.4, rect.y1 + 0.4)
    page.add_redact_annot(redaction_rect, fill=(1, 1, 1), cross_out=False)


def note_rect_candidate(page: fitz.Page) -> fitz.Rect:
    rect = page.rect
    margin_x = 18.0
    note_height = 6.0
    y1 = rect.y1 - 2.0
    y0 = y1 - note_height
    return fitz.Rect(rect.x0 + margin_x, y0, rect.x1 - margin_x, y1)


def has_space_for_note(page: fitz.Page, rect: fitz.Rect) -> bool:
    text_dict = page.get_text("dict")
    for block in text_dict.get("blocks", []):
        block_bbox = block.get("bbox")
        if not isinstance(block_bbox, (list, tuple)) or len(block_bbox) != 4:
            continue
        existing = fitz.Rect(float(block_bbox[0]), float(block_bbox[1]), float(block_bbox[2]), float(block_bbox[3]))
        if existing.intersects(rect):
            return False
    return True


def add_translation_note(page: fitz.Page) -> bool:
    rect = note_rect_candidate(page)
    if not has_space_for_note(page, rect):
        LOGGER.warning("Skipping translation note due to insufficient free space | page=%s", page.number + 1)
        return False

    result = page.insert_textbox(
        rect,
        TRANSLATION_NOTE,
        fontsize=4.0,
        fontname="helv",
        color=fitz.sRGB_to_pdf(0x555555),
        align=fitz.TEXT_ALIGN_CENTER,
    )
    if result < 0:
        LOGGER.warning("Failed to place translation note | page=%s", page.number + 1)
        return False
    return True


def ensure_output_paths(project_root: Path, output_path: Path | None, source_pdf: Path) -> tuple[Path, Path]:
    output_dir = project_root / "output"
    renders_dir = project_root / "renders"
    output_dir.mkdir(parents=True, exist_ok=True)
    renders_dir.mkdir(parents=True, exist_ok=True)

    if output_path is None:
        final_pdf_path = output_dir / f"{source_pdf.stem}_EN_TRANSLATION.pdf"
    else:
        final_pdf_path = output_path if output_path.is_absolute() else (project_root / output_path)

    final_pdf_path.parent.mkdir(parents=True, exist_ok=True)
    render_root = renders_dir / final_pdf_path.stem
    render_root.mkdir(parents=True, exist_ok=True)
    return final_pdf_path.resolve(), render_root.resolve()


def group_rows_by_page(rows: list[TranslationUnitRow]) -> dict[int, list[TranslationUnitRow]]:
    grouped: dict[int, list[TranslationUnitRow]] = {}
    for row in rows:
        grouped.setdefault(row.page, []).append(row)
    for page_rows in grouped.values():
        page_rows.sort(key=lambda row: (row.bbox_y0, row.bbox_x0, row.unit_id))
    return grouped


def process_pdf(
    doc: fitz.Document,
    rows: list[TranslationUnitRow],
    summary: Summary,
) -> None:
    arial_paths = {
        "arial": Path(r"C:\Windows\Fonts\arial.ttf"),
        "arialbd": Path(r"C:\Windows\Fonts\arialbd.ttf"),
        "arialbi": Path(r"C:\Windows\Fonts\arialbi.ttf"),
        "ariali": Path(r"C:\Windows\Fonts\ariali.ttf"),
    }

    rows_by_page = group_rows_by_page(rows)

    for page_index in range(doc.page_count):
        page_number = page_index + 1
        page = doc.load_page(page_index)
        page_rows = rows_by_page.get(page_number, [])
        translatable_rows: list[TranslationUnitRow] = []

        for row in page_rows:
            try:
                if row.action == "KEEP":
                    summary.keep_ignored += 1
                    continue
                if row.action not in {"TRANSLATE", "REMOVE"}:
                    continue

                if row.page < 1 or row.page > doc.page_count:
                    raise ValueError("Page index out of range")

                add_redaction(page, row)
                if row.action == "REMOVE":
                    summary.remove_processed += 1
                    continue

                if row.english_translation.strip() == "":
                    summary.empty_translations += 1
                    LOGGER.warning(
                        "Empty translation; skipping insertion | unit_id=%s | page=%s | original=%s",
                        row.unit_id,
                        row.page,
                        row.original_text,
                    )
                    continue

                translatable_rows.append(row)
                summary.translate_processed += 1

            except Exception as exc:  # noqa: BLE001
                summary.row_failures += 1
                LOGGER.error(
                    "Row processing failed before redaction apply | unit_id=%s | page=%s | text=%s | error=%s",
                    row.unit_id,
                    row.page,
                    row.original_text,
                    exc,
                )

        if page.first_annot is not None:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE)

        for row in translatable_rows:
            try:
                font_spec = choose_font_spec(row, arial_paths)
                try_insert_translation(page, row, font_spec, summary)
            except Exception as exc:  # noqa: BLE001
                summary.row_failures += 1
                LOGGER.error(
                    "Row insertion failed | unit_id=%s | page=%s | text=%s | error=%s",
                    row.unit_id,
                    row.page,
                    row.original_text,
                    exc,
                )

        add_translation_note(page)


def render_validation_images(doc_path: Path, render_root: Path) -> None:
    dpi = 150
    zoom = dpi / 72.0
    matrix = fitz.Matrix(zoom, zoom)
    with fitz.open(doc_path) as doc:
        for index in range(doc.page_count):
            page = doc.load_page(index)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            pix.save(render_root / f"page_{index + 1:03d}.png")


def print_summary(summary: Summary, output_pdf: Path, render_root: Path) -> None:
    print("Translation apply summary")
    print(f"- TRANSLATE units processed: {summary.translate_processed}")
    print(f"- REMOVE units processed: {summary.remove_processed}")
    print(f"- KEEP units ignored: {summary.keep_ignored}")
    print(f"- Empty translations: {summary.empty_translations}")
    print(f"- Texts with reduced font: {summary.reduced_font_texts}")
    print(f"- Texts that did not fit: {summary.not_fit_texts}")
    print(f"- Row failures: {summary.row_failures}")
    print(f"- Output PDF: {output_pdf}")
    print(f"- Validation renders: {render_root}")


def main() -> int:
    setup_logging()
    args = parse_args()

    script_path = Path(__file__).resolve()
    project_root = script_path.parent.parent
    pdf_path = to_project_path(project_root, args.pdf_path)
    excel_path = to_project_path(project_root, args.excel_path)

    if not pdf_path.exists():
        LOGGER.error("PDF not found: %s", pdf_path)
        return 1
    if not excel_path.exists():
        LOGGER.error("Excel not found: %s", excel_path)
        return 1

    output_pdf, render_root = ensure_output_paths(project_root, args.output, pdf_path)

    try:
        rows = read_translation_units(excel_path, args.sheet)
    except Exception as exc:  # noqa: BLE001
        LOGGER.error("Failed to read translation sheet: %s", exc)
        return 1

    try:
        with fitz.open(pdf_path) as doc:
            summary = Summary()
            process_pdf(doc, rows, summary)
            doc.save(output_pdf, garbage=4, deflate=True, clean=True)
    except Exception as exc:  # noqa: BLE001
        LOGGER.error("Failed to process PDF: %s", exc)
        return 1

    try:
        render_validation_images(output_pdf, render_root)
    except Exception as exc:  # noqa: BLE001
        LOGGER.error("Failed to render validation images: %s", exc)
        return 1

    print_summary(summary, output_pdf, render_root)
    return 0


if __name__ == "__main__":
    sys.exit(main())
