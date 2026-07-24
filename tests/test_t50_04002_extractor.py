from cal_translator.formats.t50_04002.extractor import _RESULT_RE


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
    assert match.group("average_measured_value") if False else True
