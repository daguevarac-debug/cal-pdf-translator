from cal_translator.formats.t50_04002.extractor import (
    _RESULT_RE,
    _extract_notes,
    _extract_results,
)
from cal_translator.pdf.text_extractor import PdfPageText


def test_result_row_with_range() -> None:
    match = _RESULT_RE.match(
        "1,0 Ω / 45 A 0,100 mΩ 0,101 mΩ 0,0007 mΩ 0,0058 mΩ 2,0"
    )
    assert match is not None
    assert match.group("range") == "1,0 Ω / 45 A"
    assert match.group("specified") == "0,100 mΩ"
    assert match.group("coverage") == "2,0"


def test_result_row_without_range() -> None:
    match = _RESULT_RE.match("1,000 mΩ 1,007 mΩ 0,0072 mΩ 0,0058 mΩ 2,0")
    assert match is not None
    assert match.group("range") is None
    assert match.group("measured") == "1,007 mΩ"
    assert match.group("bias") == "0,0072 mΩ"


def test_results_are_partitioned_into_three_24_row_channels() -> None:
    first = "1,0 Ω / 45 A 0,100 mΩ 0,101 mΩ 0,0007 mΩ 0,0058 mΩ 2,0"
    continuation = "1,000 mΩ 1,007 mΩ 0,0072 mΩ 0,0058 mΩ 2,0"
    page = PdfPageText(page_number=3, text="\n".join([first, *([continuation] * 71)]))

    rows = _extract_results([page])

    assert len(rows) == 72
    assert [sum(row.channel == channel for row in rows) for channel in "ABC"] == [24, 24, 24]
    assert rows[0].range == "1,0 Ω / 45 A"
    assert rows[-1].range == "1,0 Ω / 45 A"


def test_notes_are_sorted_by_source_number() -> None:
    page = PdfPageText(
        page_number=5,
        text="7. Seventh note\ncontinued\n2. Second note\n** Fin del certificado **",
    )

    assert _extract_notes(page) == ["2. Second note", "7. Seventh note continued"]
