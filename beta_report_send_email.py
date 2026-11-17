"""Legacy filename retained for migration. Generates files only; never sends email.

Copyright 2018 Kleber de Oliveira Andrade.
"""

import sys
from class_report.cli import main

if __name__ == "__main__":
    raise SystemExit(main([*(sys.argv[1:] or ["beta evaluation.xlsx"]), "--stage", "beta"]))
