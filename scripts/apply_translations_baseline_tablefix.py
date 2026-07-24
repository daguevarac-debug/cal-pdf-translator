from __future__ import annotations

import statistics
import sys

import apply_translations as core
import apply_translations_baseline as baseline


_ORIGINAL_BUILD_BASELINE_PLANS = baseline._build_baseline_plans


def _build_baseline_plans_with_table_header_size(
    doc: core.fitz.Document,
    rows: list[core.TranslationUnitRow],
    span_map: dict[int, core.TextSpanRow],
) -> None:
    """Reuse the page-2 table header group size for split header pairs."""

    _ORIGINAL_BUILD_BASELINE_PLANS(doc, rows, span_map)

    rows_by_page = core.group_rows_by_page(rows)
    arial_paths = core.collect_arial_paths()

    for page_index in range(doc.page_count):
        page_number = page_index + 1
        if page_number != 2:
            continue

        page = doc.load_page(page_index)
        page_rows = rows_by_page.get(page_number, [])
        specs, _ = core.build_page_specs(page.rect, page_rows, span_map, arial_paths)

        grouped_header_sizes = [
            specs[row.unit_id].font_size_final
            for row in page_rows
            if row.unit_id in specs
            and specs[row.unit_id].placement_mode == "TABLE_GROUP"
            and core.is_page2_table_header_row(row)
        ]
        if not grouped_header_sizes:
            continue

        common_header_size = round(statistics.median(grouped_header_sizes), 2)
        pairs = baseline._find_table_pairs(page_rows, span_map)

        for first_id, (first, second, combined_text) in pairs.items():
            first_origins, _ = baseline._source_lines(first, span_map)
            second_origins, _ = baseline._source_lines(second, span_map)
            origins = [first_origins[0], second_origins[0]]
            bbox = core.fitz.Rect(
                min(first.bbox_x0, second.bbox_x0),
                min(first.bbox_y0, second.bbox_y0),
                max(first.bbox_x1, second.bbox_x1),
                max(first.bbox_y1, second.bbox_y1),
            )

            baseline._BASELINE_PLANS[first_id] = baseline._plan_baseline(
                first,
                combined_text,
                common_header_size,
                page.rect,
                page_rows,
                span_map,
                forced_alignment="CENTER",
                forced_bbox=bbox,
                forced_origins=origins,
                report_mode="TABLE_GROUP",
            )

            baseline._BASELINE_PLANS[second.unit_id] = baseline.BaselinePlan(
                fits=True,
                alignment="CENTER",
                insertion_x=origins[-1][0],
                insertion_y=origins[-1][1],
                available_left=bbox.x0,
                available_right=bbox.x1,
                available_width=bbox.width,
                required_width=0.0,
                font_size=common_header_size,
                rendered_lines=[],
                line_positions=[],
                used_real_origin=True,
                consumed=True,
                report_mode="TABLE_GROUP",
            )


def main() -> int:
    baseline._build_baseline_plans = _build_baseline_plans_with_table_header_size
    return baseline.main()


if __name__ == "__main__":
    sys.exit(main())
