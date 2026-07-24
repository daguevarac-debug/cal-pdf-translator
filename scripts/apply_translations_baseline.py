from __future__ import annotations

import itertools
import re
import statistics
import sys
from dataclasses import dataclass
from typing import Any

import apply_translations as core


@dataclass(slots=True)
class BaselinePlan:
    fits: bool
    alignment: str
    insertion_x: float
    insertion_y: float
    available_left: float
    available_right: float
    available_width: float
    required_width: float
    font_size: float
    rendered_lines: list[str]
    line_positions: list[tuple[float, float]]
    used_real_origin: bool
    consumed: bool = False
    report_mode: str = "BASELINE"


_BASELINE_PLANS: dict[int, BaselinePlan] = {}
_ORIGINAL_RUN_PREFLIGHT = core.run_preflight
_ORIGINAL_TRY_INSERT_TRANSLATION = core.try_insert_translation


def _normalized(value: str) -> str:
    return core.normalize_text(value).strip("* ")


def _alignment_name(row: core.TranslationUnitRow, page_rect: core.fitz.Rect) -> str:
    text = _normalized(core.display_text_for_row(row))
    if text in {"calibration certificate", "ratio measurement", "end of certificate"}:
        return "CENTER"
    if text == "<restricted>":
        return "RIGHT"
    if re.fullmatch(r"page\s+\d+\s+of\s+\d+", text):
        return "RIGHT"
    if row.bbox_x1 >= page_rect.x1 - 20.0:
        return "RIGHT"
    return "LEFT"


def _row_spans(row: core.TranslationUnitRow, span_map: dict[int, core.TextSpanRow]) -> list[core.TextSpanRow]:
    return core.get_unit_spans(row, span_map)


def _source_lines(
    row: core.TranslationUnitRow,
    span_map: dict[int, core.TextSpanRow],
) -> tuple[list[tuple[float, float]], bool]:
    spans = _row_spans(row, span_map)
    if not spans:
        return [(row.bbox_x0, row.bbox_y1)], False

    grouped: dict[tuple[int, int], list[core.TextSpanRow]] = {}
    for span in spans:
        grouped.setdefault((span.block_number, span.line_number), []).append(span)

    origins: list[tuple[float, float]] = []
    for line_spans in grouped.values():
        origins.append(
            (
                min(span.origin_x for span in line_spans),
                statistics.median(span.origin_y for span in line_spans),
            )
        )
    origins.sort(key=lambda item: (item[1], item[0]))
    return origins or [(row.bbox_x0, row.bbox_y1)], True


def _text_width(text: str, row: core.TranslationUnitRow, font_size: float) -> float:
    font_spec = core.choose_font_spec(row, core.collect_arial_paths())
    return core.estimate_text_width(text, font_spec, font_size)


def _split_words_balanced(
    text: str,
    line_count: int,
    row: core.TranslationUnitRow,
    font_size: float,
) -> list[str]:
    words = text.split()
    if line_count <= 1 or len(words) <= 1:
        return [text]
    line_count = min(line_count, len(words))

    best_lines: list[str] | None = None
    best_score: tuple[float, float] | None = None
    split_slots = range(1, len(words))
    for cuts in itertools.combinations(split_slots, line_count - 1):
        indices = (0, *cuts, len(words))
        lines = [" ".join(words[indices[i] : indices[i + 1]]) for i in range(line_count)]
        widths = [_text_width(line, row, font_size) for line in lines]
        score = (max(widths), max(widths) - min(widths))
        if best_score is None or score < best_score:
            best_score = score
            best_lines = lines
    return best_lines or [text]


def _rendered_lines(
    row: core.TranslationUnitRow,
    text: str,
    source_origins: list[tuple[float, float]],
    font_size: float,
) -> list[str]:
    explicit = text.splitlines()
    if len(explicit) > 1:
        return explicit
    if len(source_origins) > 1:
        return _split_words_balanced(text, len(source_origins), row, font_size)
    return [text]


def _reference_y(row: core.TranslationUnitRow, span_map: dict[int, core.TextSpanRow]) -> float:
    origins, _ = _source_lines(row, span_map)
    return origins[0][1]


def _same_visual_row(
    target_y: float,
    target: core.TranslationUnitRow,
    other: core.TranslationUnitRow,
    span_map: dict[int, core.TextSpanRow],
) -> bool:
    other_y = _reference_y(other, span_map)
    tolerance = max(target.font_size, other.font_size, 1.0) * 0.45
    if abs(target_y - other_y) <= tolerance:
        return True

    overlap = min(target.bbox_y1, other.bbox_y1) - max(target.bbox_y0, other.bbox_y0)
    min_height = max(0.1, min(target.bbox_y1 - target.bbox_y0, other.bbox_y1 - other.bbox_y0))
    return overlap >= min_height * 0.5


def _line_bounds(
    row: core.TranslationUnitRow,
    line_y: float,
    alignment: str,
    page_rect: core.fitz.Rect,
    page_rows: list[core.TranslationUnitRow],
    span_map: dict[int, core.TextSpanRow],
) -> tuple[float, float]:
    left = page_rect.x0 + 10.0
    right = page_rect.x1 - 10.0

    obstacles = [
        other
        for other in page_rows
        if other.unit_id != row.unit_id
        and other.action in {"TRANSLATE", "REMOVE", "KEEP"}
        and _same_visual_row(line_y, row, other, span_map)
    ]

    left_candidates = [
        other.bbox_x1 + 3.0
        for other in obstacles
        if other.bbox_x1 <= row.bbox_x0 + 0.5
    ]
    right_candidates = [
        other.bbox_x0 - 3.0
        for other in obstacles
        if other.bbox_x0 >= row.bbox_x1 - 0.5
    ]

    if left_candidates:
        left = max(left, max(left_candidates))
    if right_candidates:
        right = min(right, min(right_candidates))

    if alignment == "LEFT":
        left = max(left, row.bbox_x0 - 0.5)
    elif alignment == "RIGHT":
        right = min(right, row.bbox_x1 + 0.5)

    return left, right


def _plan_baseline(
    row: core.TranslationUnitRow,
    text: str,
    font_size: float,
    page_rect: core.fitz.Rect,
    page_rows: list[core.TranslationUnitRow],
    span_map: dict[int, core.TextSpanRow],
    *,
    forced_alignment: str | None = None,
    forced_bbox: core.fitz.Rect | None = None,
    forced_origins: list[tuple[float, float]] | None = None,
    report_mode: str = "BASELINE",
) -> BaselinePlan:
    source_origins, used_real_origin = _source_lines(row, span_map)
    if forced_origins is not None:
        source_origins = forced_origins
        used_real_origin = True

    alignment = forced_alignment or _alignment_name(row, page_rect)
    rendered_lines = _rendered_lines(row, text, source_origins, font_size)

    if len(source_origins) > 1:
        positive = [
            source_origins[i + 1][1] - source_origins[i][1]
            for i in range(len(source_origins) - 1)
            if source_origins[i + 1][1] > source_origins[i][1]
        ]
        spacing = statistics.median(positive) if positive else font_size * 1.2
    else:
        spacing = font_size * 1.2

    positions: list[tuple[float, float]] = []
    available_widths: list[float] = []
    required_widths: list[float] = []
    fit_flags: list[bool] = []
    first_left = page_rect.x0 + 10.0
    first_right = page_rect.x1 - 10.0

    target_bbox = forced_bbox or row.bbox
    target_center = (target_bbox.x0 + target_bbox.x1) / 2.0

    for index, line in enumerate(rendered_lines):
        if index < len(source_origins):
            source_x, line_y = source_origins[index]
        else:
            source_x = source_origins[-1][0]
            line_y = source_origins[-1][1] + spacing * (index - len(source_origins) + 1)

        left, right = _line_bounds(row, line_y, alignment, page_rect, page_rows, span_map)
        if forced_bbox is not None:
            left = max(left, forced_bbox.x0)
            right = min(right, forced_bbox.x1)

        required = _text_width(line, row, font_size)
        if alignment == "RIGHT":
            insertion_x = min(target_bbox.x1, right) - required
        elif alignment == "CENTER":
            insertion_x = target_center - required / 2.0
        else:
            insertion_x = max(source_x, left)

        fits = insertion_x >= left - 0.1 and insertion_x + required <= right + 0.1
        positions.append((insertion_x, line_y))
        available_widths.append(max(0.0, right - left))
        required_widths.append(required)
        fit_flags.append(fits)
        if index == 0:
            first_left, first_right = left, right

    return BaselinePlan(
        fits=all(fit_flags),
        alignment=alignment,
        insertion_x=positions[0][0],
        insertion_y=positions[0][1],
        available_left=first_left,
        available_right=first_right,
        available_width=min(available_widths, default=0.0),
        required_width=max(required_widths, default=0.0),
        font_size=font_size,
        rendered_lines=rendered_lines,
        line_positions=positions,
        used_real_origin=used_real_origin,
        report_mode=report_mode,
    )


def _find_table_pairs(
    page_rows: list[core.TranslationUnitRow],
    span_map: dict[int, core.TextSpanRow],
) -> dict[int, tuple[core.TranslationUnitRow, core.TranslationUnitRow, str]]:
    if not page_rows or page_rows[0].page != 2:
        return {}

    by_text: dict[str, list[core.TranslationUnitRow]] = {}
    for row in page_rows:
        if row.action != "TRANSLATE":
            continue
        by_text.setdefault(_normalized(core.display_text_for_row(row)), []).append(row)

    pairs: dict[int, tuple[core.TranslationUnitRow, core.TranslationUnitRow, str]] = {}
    for first_text, second_text, combined in (
        ("internal", "number", "Internal\nNumber"),
        ("calibration", "date", "Calibration\nDate"),
    ):
        first_rows = by_text.get(first_text, [])
        second_rows = by_text.get(second_text, [])
        for first in first_rows:
            candidates = []
            for second in second_rows:
                horizontal_overlap = min(first.bbox_x1, second.bbox_x1) - max(first.bbox_x0, second.bbox_x0)
                if horizontal_overlap <= 0:
                    continue
                dy = abs(_reference_y(first, span_map) - _reference_y(second, span_map))
                if dy <= max(first.font_size, second.font_size) * 2.5:
                    candidates.append((dy, second))
            if candidates:
                _, second = min(candidates, key=lambda item: item[0])
                pairs[first.unit_id] = (first, second, combined)
                break
    return pairs


def _build_baseline_plans(
    doc: core.fitz.Document,
    rows: list[core.TranslationUnitRow],
    span_map: dict[int, core.TextSpanRow],
) -> None:
    _BASELINE_PLANS.clear()
    rows_by_page = core.group_rows_by_page(rows)
    arial_paths = core.collect_arial_paths()

    for page_index in range(doc.page_count):
        page_number = page_index + 1
        page = doc.load_page(page_index)
        page_rows = rows_by_page.get(page_number, [])
        specs, _ = core.build_page_specs(page.rect, page_rows, span_map, arial_paths)

        pairs = _find_table_pairs(page_rows, span_map)
        consumed_ids: set[int] = set()
        for first_id, (first, second, combined) in pairs.items():
            first_origins, _ = _source_lines(first, span_map)
            second_origins, _ = _source_lines(second, span_map)
            origins = [first_origins[0], second_origins[0]]
            bbox = core.fitz.Rect(
                min(first.bbox_x0, second.bbox_x0),
                min(first.bbox_y0, second.bbox_y0),
                max(first.bbox_x1, second.bbox_x1),
                max(first.bbox_y1, second.bbox_y1),
            )
            spec = specs.get(first_id, core.build_default_spec(first))
            _BASELINE_PLANS[first_id] = _plan_baseline(
                first,
                combined,
                spec.font_size_final,
                page.rect,
                page_rows,
                span_map,
                forced_alignment="CENTER",
                forced_bbox=bbox,
                forced_origins=origins,
                report_mode="TABLE_GROUP",
            )
            _BASELINE_PLANS[second.unit_id] = BaselinePlan(
                fits=True,
                alignment="CENTER",
                insertion_x=origins[-1][0],
                insertion_y=origins[-1][1],
                available_left=bbox.x0,
                available_right=bbox.x1,
                available_width=bbox.width,
                required_width=0.0,
                font_size=spec.font_size_final,
                rendered_lines=[],
                line_positions=[],
                used_real_origin=True,
                consumed=True,
                report_mode="TABLE_GROUP",
            )
            consumed_ids.add(second.unit_id)

        for row in page_rows:
            spec = specs.get(row.unit_id)
            if row.action != "TRANSLATE" or spec is None:
                continue
            if row.unit_id in _BASELINE_PLANS or row.unit_id in consumed_ids:
                continue
            if spec.placement_mode != "BASELINE":
                continue
            text = spec.forced_text if spec.forced_text is not None else row.english_translation
            _BASELINE_PLANS[row.unit_id] = _plan_baseline(
                row,
                text,
                spec.font_size_final,
                page.rect,
                page_rows,
                span_map,
            )


def run_preflight(
    doc: core.fitz.Document,
    rows: list[core.TranslationUnitRow],
    span_map: dict[int, core.TextSpanRow],
) -> dict[str, Any]:
    report = _ORIGINAL_RUN_PREFLIGHT(doc, rows, span_map)
    _build_baseline_plans(doc, rows, span_map)

    entries_by_id = {int(entry["unit_id"]): entry for entry in report.get("units", [])}
    original_baseline_failed_pages = {
        int(entry["page"])
        for entry in report.get("units", [])
        if entry.get("placement_mode") == "BASELINE" and not entry.get("fits", False)
    }
    preserved_failed_pages = set(report.get("failed_pages", [])) - original_baseline_failed_pages

    for unit_id, plan in _BASELINE_PLANS.items():
        entry = entries_by_id.get(unit_id)
        if entry is None:
            continue
        entry.update(
            {
                "fits": plan.fits,
                "placement_mode": plan.report_mode,
                "alignment": plan.alignment,
                "available_width": round(plan.available_width, 2),
                "required_width": round(plan.required_width, 2),
                "insertion_x": round(plan.insertion_x, 2),
                "insertion_y": round(plan.insertion_y, 2),
                "used_real_origin": plan.used_real_origin,
                "rendered_lines": plan.rendered_lines,
                "consumed_by_group": plan.consumed,
            }
        )

    failed_pages = set(preserved_failed_pages)
    for entry in report.get("units", []):
        if not entry.get("fits", False):
            failed_pages.add(int(entry["page"]))

    report["failed_pages"] = sorted(failed_pages)
    report["status"] = "FAILED_PREFLIGHT" if failed_pages else "OK"
    report["page_status"] = {
        str(page_number): ("FAILED_PREFLIGHT" if page_number in failed_pages else "OK")
        for page_number in range(1, doc.page_count + 1)
    }
    return report


def try_insert_translation(
    page: core.fitz.Page,
    unit: core.TranslationUnitRow,
    font_spec: core.FontSpec,
    summary: core.Summary,
    spec: core.PlacementSpec,
    spans: list[core.TextSpanRow],
) -> bool:
    plan = _BASELINE_PLANS.get(unit.unit_id)
    if plan is None:
        return _ORIGINAL_TRY_INSERT_TRANSLATION(page, unit, font_spec, summary, spec, spans)
    if plan.consumed:
        return True
    if not plan.fits:
        summary.not_fit_texts += 1
        return False

    color = core.fitz.sRGB_to_pdf(unit.color)
    for line, (x, y) in zip(plan.rendered_lines, plan.line_positions, strict=True):
        page.insert_text(
            point=(x, y),
            text=line,
            fontsize=plan.font_size,
            fontname=font_spec.fontname,
            fontfile=font_spec.fontfile,
            color=color,
            render_mode=0,
        )
    return True


def main() -> int:
    core.run_preflight = run_preflight
    core.try_insert_translation = try_insert_translation
    return core.main()


if __name__ == "__main__":
    sys.exit(main())
