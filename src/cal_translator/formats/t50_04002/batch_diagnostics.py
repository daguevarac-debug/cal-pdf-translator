from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from cal_translator.formats.t50_04002 import batch_processor as legacy

_ORIGINAL_PROCESS_ONE = legacy._process_one


def _summarize_postflight_error(item: dict[str, Any]) -> str:
    code = str(item.get("code") or "postflight")
    target = str(item.get("target") or "unknown target")
    expected = item.get("expected")
    actual = item.get("actual")
    return f"{code} at {target}: expected {expected!r}, actual {actual!r}"


def _process_one(
    certificate_number: str,
    pdf_path: Path,
    template_path: Path,
    output_root: Path,
    *,
    overwrite: bool,
) -> dict[str, Any]:
    result = _ORIGINAL_PROCESS_ONE(
        certificate_number,
        pdf_path,
        template_path,
        output_root,
        overwrite=overwrite,
    )
    if (
        result.get("status") != "failed"
        or result.get("error") != "Expanded workbook postflight validation failed."
    ):
        return result

    certificate_dir = output_root / certificate_number
    build_report_path = certificate_dir / f"{certificate_number}_WORKING_build_report.json"
    if not build_report_path.exists():
        return result

    try:
        build_report = json.loads(build_report_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return result

    errors = list(build_report.get("postflight_errors") or [])
    result["build_postflight_errors"] = errors
    if errors:
        details = "; ".join(_summarize_postflight_error(item) for item in errors[:4])
        if len(errors) > 4:
            details += f"; plus {len(errors) - 4} more"
        result["error"] = f"Expanded workbook postflight validation failed: {details}"
        legacy._write_report(
            result,
            certificate_dir / f"{certificate_number}_process_report.json",
        )
    return result


def install_patch() -> None:
    """Expose detailed workbook postflight diagnostics in batch summaries."""
    legacy._process_one = _process_one
