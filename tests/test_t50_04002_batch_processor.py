from pathlib import Path

import pytest

from cal_translator.formats.t50_04002.batch_processor import (
    certificate_id,
    certificate_range,
    find_certificate_pdf,
)


def test_certificate_range_is_inclusive() -> None:
    assert certificate_range(12513, 12521) == [
        "CAL-12513",
        "CAL-12514",
        "CAL-12515",
        "CAL-12516",
        "CAL-12517",
        "CAL-12518",
        "CAL-12519",
        "CAL-12520",
        "CAL-12521",
    ]


def test_certificate_range_rejects_reverse_bounds() -> None:
    with pytest.raises(ValueError, match="must not exceed"):
        certificate_range(12521, 12513)


def test_certificate_id_rejects_nonpositive_numbers() -> None:
    with pytest.raises(ValueError, match="positive"):
        certificate_id(0)


def test_find_certificate_pdf_accepts_expected_filename(tmp_path: Path) -> None:
    pdf_path = tmp_path / "CAL-12513.pdf"
    pdf_path.write_bytes(b"%PDF-placeholder")

    assert find_certificate_pdf(tmp_path, "CAL-12513") == pdf_path


def test_find_certificate_pdf_returns_none_when_missing(tmp_path: Path) -> None:
    assert find_certificate_pdf(tmp_path, "CAL-12513") is None
