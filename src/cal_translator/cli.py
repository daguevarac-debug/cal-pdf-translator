from __future__ import annotations

import argparse
import logging
from pathlib import Path

from cal_translator.excel.template_inventory import inspect_template, write_inventory


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="cal_translator")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")

    subparsers = parser.add_subparsers(dest="command", required=True)

    inspect_parser = subparsers.add_parser(
        "inspect-template",
        help="Inspect an Excel certificate template through Microsoft Excel COM without modifying it.",
    )
    inspect_parser.add_argument("--template", type=Path, required=True, help="Path to the source Excel template")
    inspect_parser.add_argument("--format", dest="format_id", required=True, help="Controlled document format, e.g. T50-04002")
    inspect_parser.add_argument("--output", type=Path, required=True, help="Directory for JSON and Markdown inventory reports")

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
            return 0
    except (FileNotFoundError, RuntimeError) as exc:
        parser.exit(2, f"ERROR: {exc}\n")

    parser.error(f"Unsupported command: {args.command}")
    return 2
