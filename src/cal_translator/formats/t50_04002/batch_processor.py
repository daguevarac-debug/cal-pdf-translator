from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from cal_translator.formats.t50_04002.english_export_v3 import translate_and_export
from cal_translator.formats.t50_04002.extractor import extract_certificate, write_extraction
from cal_translator.formats.t50_04002.workbook_writer_v2 import build_workbook_prototype

FORMAT_ID = "T50-04002"
_MAX_RANGE_SIZE = 500


def certificate_id(number: int) -> str:
    if number <= 0:
        raise ValueError("Certificate numbers must be positive integers.")
    return f"CAL-{number}"


def certificate_range(start: int, end: int) -> list[str]:
    if start > end:
        raise ValueError("The start certificate number must not exceed the end number.")
    count = end - start + 1
    if count > _MAX_RANGE_SIZE:
        raise ValueError(
            f"A single batch may contain at most {_MAX_RANGE_SIZE} certificates; requested {count}."
        )
    return [certificate_id(number) for number in range(start, end + 1)]


def find_certificate_pdf(input_dir: Path, expected_id: str) -> Path | None:
    direct = input_dir / f"{expected_id}.pdf"
    if direct.exists():
        return direct

    candidates = sorted(
        path
        for path in input_dir.glob("*.pdf")
        if path.stem.casefold() == expected_id.casefold()
    )
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        raise RuntimeError(
            f"Multiple PDF files match {expected_id}: "
            + ", ".join(path.name for path in candidates)
        )
    return None


def _utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _write_report(report: dict[str, Any], report_path: Path) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    report_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2, default=str),
        encoding="utf-8",
    )


def _process_one(
    certificate_number: str,
    pdf_path: Path,
    template_path: Path,
    output_root: Path,
    *,
    overwrite: bool,
) -> dict[str, Any]:
    certificate_dir = output_root / certificate_number
    certificate_dir.mkdir(parents=True, exist_ok=True)

    result: dict[str, Any] = {
        "certificate": certificate_number,
        "source_pdf": str(pdf_path.resolve()),
        "output_directory": str(certificate_dir.resolve()),
        "status": "processing",
        "started_at": _utc_now(),
    }

    try:
        extracted = extract_certificate(pdf_path)
        extraction_path = write_extraction(extracted, certificate_dir)
        result["extraction_json"] = str(extraction_path.resolve())
        result["source_pages"] = extracted.source_pages
        result["traceability_rows"] = len(extracted.traceability)
        result["result_rows"] = len(extracted.results)
        result["channel_counts"] = {
            channel: sum(row.channel == channel for row in extracted.results)
            for channel in ("A", "B", "C")
        }

        warnings = list(extracted.extraction_warnings)
        if extracted.certificate_number != certificate_number:
            warnings.append(
                "Certificate number mismatch: "
                f"expected {certificate_number}, extracted "
                f"{extracted.certificate_number or 'none'}."
            )
        result["extraction_warnings"] = warnings
        if warnings:
            raise RuntimeError(
                f"Extraction validation failed with {len(warnings)} warning(s)."
            )

        working_path = certificate_dir / f"{certificate_number}_WORKING.xlsx"
        build_report, build_report_path = build_workbook_prototype(
            template_path,
            extraction_path,
            working_path,
            format_id=FORMAT_ID,
            overwrite=overwrite,
        )
        result["working_workbook"] = str(working_path.resolve())
        result["build_report"] = str(build_report_path.resolve())
        result["build_valid"] = bool(build_report.get("postflight_valid"))
        if not result["build_valid"]:
            raise RuntimeError("Expanded workbook postflight validation failed.")

        english_workbook = certificate_dir / f"{certificate_number}_EN_TRANSLATION.xlsx"
        english_pdf = certificate_dir / f"{certificate_number}_EN_TRANSLATION.pdf"
        translation_report, translation_report_path = translate_and_export(
            working_path,
            extraction_path,
            english_workbook,
            english_pdf,
            format_id=FORMAT_ID,
            overwrite=overwrite,
        )
        result["english_workbook"] = str(english_workbook.resolve())
        result["english_pdf"] = str(english_pdf.resolve())
        result["translation_report"] = str(translation_report_path.resolve())
        result["workbook_valid"] = bool(
            translation_report.get("workbook_validation", {}).get("valid")
        )
        result["pdf_valid"] = bool(
            translation_report.get("pdf_validation", {}).get("valid")
        )
        result["pdf_pages"] = translation_report.get("pdf_validation", {}).get(
            "page_count"
        )
        if not translation_report.get("valid"):
            raise RuntimeError("English workbook or PDF validation failed.")

        result["status"] = "completed"
    except Exception as exc:  # noqa: BLE001 - batch must preserve per-file failures
        result["status"] = "failed"
        result["error_type"] = type(exc).__name__
        result["error"] = str(exc)

    result["finished_at"] = _utc_now()
    _write_report(result, certificate_dir / f"{certificate_number}_process_report.json")
    return result


def process_certificate_range(
    input_dir: Path,
    template_path: Path,
    output_dir: Path,
    *,
    start: int,
    end: int,
    format_id: str = FORMAT_ID,
    overwrite: bool = False,
    stop_on_error: bool = False,
) -> tuple[dict[str, Any], Path]:
    if format_id != FORMAT_ID:
        raise RuntimeError(
            f"Unsupported batch format: {format_id}. Expected {FORMAT_ID}."
        )
    if not input_dir.is_dir():
        raise FileNotFoundError(f"Input PDF directory not found: {input_dir}")
    if not template_path.is_file():
        raise FileNotFoundError(f"Excel template not found: {template_path}")

    expected_certificates = certificate_range(start, end)
    output_dir.mkdir(parents=True, exist_ok=True)
    report_path = output_dir / (
        f"batch_report_{expected_certificates[0]}_{expected_certificates[-1]}.json"
    )
    report: dict[str, Any] = {
        "status": "processing",
        "format_id": format_id,
        "input_directory": str(input_dir.resolve()),
        "template": str(template_path.resolve()),
        "output_directory": str(output_dir.resolve()),
        "start_certificate": expected_certificates[0],
        "end_certificate": expected_certificates[-1],
        "expected_count": len(expected_certificates),
        "overwrite": overwrite,
        "stop_on_error": stop_on_error,
        "started_at": _utc_now(),
        "certificates": [],
    }
    _write_report(report, report_path)

    for expected_id in expected_certificates:
        try:
            pdf_path = find_certificate_pdf(input_dir, expected_id)
        except Exception as exc:  # noqa: BLE001
            item = {
                "certificate": expected_id,
                "status": "failed",
                "error_type": type(exc).__name__,
                "error": str(exc),
            }
            report["certificates"].append(item)
            _write_report(report, report_path)
            if stop_on_error:
                break
            continue

        if pdf_path is None:
            item = {
                "certificate": expected_id,
                "status": "missing",
                "error": f"Expected PDF not found: {expected_id}.pdf",
            }
        else:
            item = _process_one(
                expected_id,
                pdf_path,
                template_path,
                output_dir,
                overwrite=overwrite,
            )
        report["certificates"].append(item)
        _write_report(report, report_path)
        if stop_on_error and item["status"] != "completed":
            break

    status_counts = {
        status: sum(item["status"] == status for item in report["certificates"])
        for status in ("completed", "failed", "missing")
    }
    processed_ids = {item["certificate"] for item in report["certificates"]}
    not_processed = [
        certificate
        for certificate in expected_certificates
        if certificate not in processed_ids
    ]
    report["status_counts"] = status_counts
    report["not_processed"] = not_processed
    report["finished_at"] = _utc_now()
    report["status"] = (
        "completed"
        if status_counts["completed"] == len(expected_certificates)
        else "completed_with_errors"
    )
    report["valid"] = report["status"] == "completed"
    _write_report(report, report_path)
    return report, report_path
