from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any

from cal_translator.formats.t50_04002 import temperature_english_export as legacy

_ORIGINAL_LOAD_PROFILE = legacy._load_profile
_ORIGINAL_CONFIGURE_LAYOUT = legacy._configure_layout
_ORIGINAL_TRANSLATE_AND_EXPORT = legacy.translate_temperature_and_export

_STATIC_ADDRESS_MAP = {
    "B73": "B72",
    "B74": "B73",
}


def _load_profile() -> dict[str, Any]:
    """Load the temperature profile using the template's original, unshifted rows."""
    profile = deepcopy(_ORIGINAL_LOAD_PROFILE())
    profile["profile_version"] = 2

    for item in profile.get("static_cells") or []:
        if item.get("sheet") == "Información":
            address = str(item.get("address") or "")
            if address in _STATIC_ADDRESS_MAP:
                item["address"] = _STATIC_ADDRESS_MAP[address]

    paragraphs = profile.get("static_paragraphs") or {}
    if "result_note_1" in paragraphs:
        paragraphs["result_note_1"]["address"] = "B74"
    if "result_note_2" in paragraphs:
        paragraphs["result_note_2"]["address"] = "B76"
    return profile


def _configure_layout(workbook: Any) -> dict[str, Any]:
    """Preserve the original three-page temperature certificate print areas."""
    info = legacy.retry_com_call(lambda: workbook.Worksheets("Información"))
    results = legacy.retry_com_call(lambda: workbook.Worksheets("Resultados"))
    legacy.set_com_property(
        legacy.retry_com_call(lambda: info.PageSetup),
        "PrintArea",
        "$B$1:$H$77",
    )
    result_setup = legacy.retry_com_call(lambda: results.PageSetup)
    legacy.set_com_property(result_setup, "PrintArea", "$B$1:$H$43")
    legacy.set_com_property(result_setup, "PrintTitleRows", "")
    legacy.retry_com_call(results.ResetAllPageBreaks)
    return {
        "information_print_area": "B1:H77",
        "results_print_area": "B1:H43",
    }


def translate_temperature_and_export(
    source_workbook: Path,
    data_path: Path,
    output_workbook: Path,
    output_pdf: Path,
    *,
    format_id: str = legacy.FORMAT_ID,
    overwrite: bool = False,
):
    """Export the temperature certificate without the obsolete row shift."""
    original_load_profile = legacy._load_profile
    original_configure_layout = legacy._configure_layout
    legacy._load_profile = _load_profile
    legacy._configure_layout = _configure_layout
    try:
        return _ORIGINAL_TRANSLATE_AND_EXPORT(
            source_workbook,
            data_path,
            output_workbook,
            output_pdf,
            format_id=format_id,
            overwrite=overwrite,
        )
    finally:
        legacy._load_profile = original_load_profile
        legacy._configure_layout = original_configure_layout


def install_patch() -> None:
    """Expose the unshifted temperature exporter to the batch processor."""
    legacy.translate_temperature_and_export = translate_temperature_and_export
