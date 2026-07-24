from cal_translator.formats.t50_04002.extractor import (
    _RESULT_RE,
    _extract_notes,
    _extract_results,
    _extract_traceability,
    _visual_lines,
)
from cal_translator.pdf.text_extractor import PdfPageText, PdfWord


def _words_for_lines(lines: list[str], start_y: float = 100.0) -> list[PdfWord]:
    words: list[PdfWord] = []
    for line_index, line in enumerate(lines):
        x = 40.0
        y = start_y + line_index * 12.0
        for word_index, token in enumerate(line.split()):
            width = max(8.0, len(token) * 4.0)
            words.append(
                PdfWord(
                    x0=x,
                    y0=y,
                    x1=x + width,
                    y1=y + 8.0,
                    text=token,
                    block_number=word_index,
                    line_number=0,
                    word_number=word_index,
                )
            )
            x += width + 4.0
    return words


def test_result_row_with_range() -> None:
    match = _RESULT_RE.search(
        "1,0 Ω / 45 A 0,100 mΩ 0,101 mΩ 0,0007 mΩ 0,0058 mΩ 2,0"
    )
    assert match is not None
    assert match.group("range") == "1,0 Ω / 45 A"
    assert match.group("specified") == "0,100 mΩ"
    assert match.group("coverage") == "2,0"


def test_result_row_without_range() -> None:
    match = _RESULT_RE.search("1,000 mΩ 1,007 mΩ 0,0072 mΩ 0,0058 mΩ 2,0")
    assert match is not None
    assert match.group("range") is None
    assert match.group("measured") == "1,007 mΩ"
    assert match.group("bias") == "0,0072 mΩ"


def test_visual_lines_reconstruct_table_rows_from_coordinates() -> None:
    lines = [
        "CANAL A",
        "1,0 Ω / 45 A 0,100 mΩ 0,101 mΩ 0,0007 mΩ 0,0058 mΩ 2,0",
        "1,000 mΩ 1,007 mΩ 0,0072 mΩ 0,0058 mΩ 2,0",
    ]
    page = PdfPageText(
        page_number=3,
        text="\n".join(reversed(lines)),
        words=_words_for_lines(lines),
    )

    assert _visual_lines(page) == lines
    rows = _extract_results([page])
    assert len(rows) == 2
    assert rows[0].specified_value == "0,100 mΩ"
    assert rows[1].range == "1,0 Ω / 45 A"


def test_results_are_partitioned_into_three_24_row_channels() -> None:
    first = "1,0 Ω / 45 A 0,100 mΩ 0,101 mΩ 0,0007 mΩ 0,0058 mΩ 2,0"
    continuation = "1,000 mΩ 1,007 mΩ 0,0072 mΩ 0,0058 mΩ 2,0"
    page = PdfPageText(page_number=3, text="\n".join([first, *([continuation] * 71)]))

    rows = _extract_results([page])

    assert len(rows) == 72
    assert [sum(row.channel == channel for row in rows) for channel in "ABC"] == [24, 24, 24]
    assert rows[0].range == "1,0 Ω / 45 A"
    assert rows[-1].range == "1,0 Ω / 45 A"


def test_traceability_rows_are_reconstructed_from_coordinates() -> None:
    lines = [
        "Resistencia de 0,1 mΩ AEG 40907 SET-GAT NU11608-25 2025-08-14",
        "Resistencia de 100 mΩ AEG 40508 SET-GAT NU11708-25 2025-08-14",
        "Resistencia de 10 mΩ AEG 41407 SET-GAT NU11408-25 2025-08-14",
        "Resistencia de 1 mΩ WLN 9 40803 SET-GAT NU11508-25 2025-08-14",
        "High Power Resistance Substituter DRS-900 5040136 SET-GAD NU8008-25 2025-08-12",
    ]
    page = PdfPageText(page_number=2, text="", words=_words_for_lines(lines))

    entries = _extract_traceability(page)

    assert len(entries) == 5
    assert entries[0].equipment == "Resistencia de 0,1 mΩ"
    assert entries[3].type == "WLN 9"
    assert entries[4].certificate_number == "NU8008-25"


def test_notes_are_sorted_by_source_number() -> None:
    page = PdfPageText(
        page_number=5,
        text="7. Seventh note\ncontinued\n2. Second note\n** Fin del certificado **",
    )

    assert _extract_notes(page) == ["2. Second note", "7. Seventh note continued"]
