"""Command-line boundary for local workbook conversion."""

import argparse
from pathlib import Path
from zipfile import BadZipFile

from openpyxl.utils.exceptions import InvalidFileException

from .data import load_evaluations
from .report import generate_reports


def main(argv=None):
    parser = argparse.ArgumentParser(description="Generate local academic DOCX reports from XLSX.")
    parser.add_argument("input", type=Path, help="Google Forms .xlsx export")
    parser.add_argument("--stage", choices=("alpha", "beta"), default="alpha")
    parser.add_argument("--sheet", help="Worksheet name (default: active worksheet)")
    parser.add_argument("-o", "--output", type=Path, default=Path("reports"))
    args = parser.parse_args(argv)
    try:
        items = load_evaluations(args.input, args.stage, args.sheet)
        paths = generate_reports(items, args.output)
    except (OSError, ValueError, BadZipFile, InvalidFileException) as error:
        parser.exit(2, f"error: {error}\n")
    print(f"{len(paths)} report(s) generated in {args.output}")
    return 0
