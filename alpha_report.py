"""Compatibility entry point. Copyright 2018 Kleber de Oliveira Andrade."""

import sys
from class_report.cli import main

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:] or ["alpha evaluation.xlsx"]))
