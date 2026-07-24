from __future__ import annotations

import argparse
import logging
from pathlib import Path

from cal_translator.excel.template_inventory_v3 import inspect_template, write_inventory
from cal_translator.formats.t50_04002.english_export_v3 import translate_and_export
from cal_translator.formats.t50_04002.extractor import extract_certificate, write_extraction
from cal_translator.formats.t50_04002.validate_template import (
    validate_template,
    write_validation_report,
)
from cal_translator.formats.t50_04002.workbook_writer_v2 import build_workbook_prototype


def _add_template_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--template", type=Path, required=True, help="Path to the source Excel template")
    parser.add_argument(
        "--format",
        dest="format_id",
        required=True,
        help="Controlled document format, e.g. T50-04002",
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Directory for JSON and Markdown reports",
    )


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="cal_translator")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")

    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser(
        "inspect-template",
        help="Inspect an Excel certificate template through Microsoft Excel COM without modifying it.",
    )
    _add_template_arguments(inspect_parser)

    validate_parser = subparsers.add_parser(
        "validate-template",
        help="Validate an Excel template against its controlled format contract without modifying it.",
    )
    _add_template_arguments(validate_parser)

    extract_parser = subparsers.add_parser(
        "extract-certificate",
        help="Extract a digital calibration certificate PDF into a structured JSON document.",
    )
    extract_parser.add_argument("--pdf", type=Path, required=True, help="Path to the source certificate PDF")
    extract_parser.add_argument(
        "--format",
        dest="format_id",
        required=True,
        help="Controlled document format, currently T50-04002",
    )
    extract_parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Directory for the structured extraction JSON",
    )

    build_parser = subparsers.add_parser(
        "build-workbook-prototype",
        help=(
            "Copy a validated Excel template, write the extracted certificate data and expand "
            "the controlled traceability and result sections. Translation remains a separate stage."
        ),
    )
    build_parser.add_argument("--template", type=Path, required=True, help="Path to the source Excel template")
    build_parser.add_argument("--data", type=Path, required=True, help="Path to the extracted certificate JSON")
    build_parser.add_argument(
        "--format",
        dest="format_id",
        required=True,
        help="Controlled document format, currently T50-04002",
    )
    build_parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Output .xlsx path for the expanded working workbook",
    )
    build_parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace an existing output workbook",
    )

    translate_parser = subparsers.add_parser(
        "translate-export",
        help=(
            "Create an American English copy of an expanded T50-04002 workbook and export "
            "the certificate sheets to PDF through Microsoft Excel."
        ),
    )
    translate_parser.add_argument(
        "--workbook",
        type=Path,
        required=True,
        help="Path to the validated expanded workbook",
    )
    translate_parser.add_argument(
        "--data",
        type=Path,
        required=True,
        help="Path to the extracted certificate JSON used to build the workbook",
    )
    translate_parser.add_argument(
        "--format",
        dest="format_id",
        required=True,
        help="Controlled document format, currently T50-04002",
    )
    translate_parser.add_argument(
        "--output-workbook",
        type=Path,
        required=True,
        help="Output path for the English .xlsx workbook",
    )
    translate_parser.add_argument(
        "--output-pdf",
        type=Path,
        required=True,
        help="Output path for the English certificate PDF",
    )
    translate_parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace existing English workbook and PDF outputs",
    )

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )

    try:
        if args.command == "inspect-template":
            inventory = inspect_template(args.template, args.format_id)
            json_path, markdown_path = write_inventory(inventory, args.output)
            print("Template inspection completed")
            print(f"- JSON: {json_path.resolve()}")
            print(f"- Markdown: {markdown_path.resolve()}")
            print(f"- Worksheets: {inventory['template']['worksheets_count']}")
            print(f"- Inventory schema: {inventory['schema_version']}")
            return 0

        if args.command == "validate-template":
            report = validate_template(args.template, args.format_id)
            json_path, markdown_path = write_validation_report(report, args.output)
            print("Template validation completed")
            print(f"- JSON: {json_path.resolve()}")
            print(f"- Markdown: {markdown_path.resolve()}")
            print(f"- Valid: {str(report['valid']).lower()}")
            print(f"- Errors: {len(report['errors'])}")
            print(f"- Warnings: {len(report['warnings'])}")
            print(f"- Checks: {len(report['checks'])}")
            return 0 if report["valid"] else 1

        if args.command == "extract-certificate":
            if args.format_id != "T50-04002":
                raise RuntimeError(
                    f"Unsupported extraction format: {args.format_id}. Expected T50-04002."
                )
            certificate = extract_certificate(args.pdf)
            json_path = write_extraction(certificate, args.output)
            channel_counts = {
                channel: sum(row.channel == channel for row in certificate.results)
                for channel in ("A", "B", "C")
            }
            print("Certificate extraction completed")
            print(f"- JSON: {json_path.resolve()}")
            print(f"- Certificate: {certificate.certificate_number or 'not detected'}")
            print(f"- Pages: {certificate.source_pages}")
            print(f"- Traceability entries: {len(certificate.traceability)}")
            print(f"- Result rows: {len(certificate.results)}")
            print(
                "- Channels: "
                + ", ".join(f"{channel}={count}" for channel, count in channel_counts.items())
            )
            print(f"- Warnings: {len(certificate.extraction_warnings)}")
            for warning in certificate.extraction_warnings:
                print(f"  - {warning}")
            return 0 if not certificate.extraction_warnings else 1

        if args.command == "build-workbook-prototype":
            report, report_path = build_workbook_prototype(
                args.template,
                args.data,
                args.output,
                format_id=args.format_id,
                overwrite=args.overwrite,
            )
            print("Workbook expansion completed")
            print(f"- Workbook: {Path(report['output_workbook']).resolve()}")
            print(f"- Report: {report_path.resolve()}")
            print(f"- Scalar fields written: {report['scalar_fields_written']}")
            print(f"- Calibration method written: {str(report['method_written']).lower()}")
            print(f"- Traceability rows written: {report['traceability_rows_written']}")
            print(f"- Result rows written: {report['result_rows_written']}")
            print(f"- Notes written: {report['notes_written']}")
            print(f"- Postflight valid: {str(report['postflight_valid']).lower()}")
            return 0 if report["postflight_valid"] else 1

        if args.command == "translate-export":
            report, report_path = translate_and_export(
                args.workbook,
                args.data,
                args.output_workbook,
                args.output_pdf,
                format_id=args.format_id,
                overwrite=args.overwrite,
            )
            pdf_validation = report["pdf_validation"]
            print("English translation and PDF export completed")
            print(f"- Workbook: {Path(report['output_workbook']).resolve()}")
            print(f"- PDF: {Path(report['output_pdf']).resolve()}")
            print(f"- Report: {report_path.resolve()}")
            print(f"- Locale: {report['locale']}")
            print(f"- Translation operations: {report['translation_operations']}")
            print(f"- PDF pages: {pdf_validation['page_count']}")
            print(f"- Workbook valid: {str(report['workbook_validation']['valid']).lower()}")
            print(f"- PDF valid: {str(pdf_validation['valid']).lower()}")
            return 0 if report["valid"] else 1
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        parser.exit(2, f"ERROR: {exc}\n")

    parser.error(f"Unsupported command: {args.command}")
    return 2
