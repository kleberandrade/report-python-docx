# Verification record

Verified locally on 17 November 2026 on macOS with Python 3.13.18.

- `pytest`: 27 behavior tests passed. Coverage measured 263 of 265 statements (99% rounded). `__main__.py` was exercised in a subprocess, so its two statements are absent from the in-process coverage report. No coverage badge is asserted.
- Ruff 0.14.5 lint and formatting checks passed for the package, tests, and compatibility wrappers.
- `python -m build --no-isolation`: source archive and universal Python wheel built successfully with setuptools 80.9.0. Installed the wheel with `--no-deps --force-reinstall` and generated both alpha and beta reports using the installed `class-report` command.
- Original alpha workbook: validated 136 response rows across nine games, without copying the historical responses into public examples. The original beta workbook is not present, so beta compatibility was verified against its script-defined schema and synthetic inputs only.
- `crapper --source-root src --use-existing-coverage --threshold 10`: passed, worst function CRAP 7.0. An initial attempt used a coverage output path the tool did not discover; reran with `target/coverage/python/lcov.info`, rather than treating missing coverage as actual zero coverage.
- `dryer --source-root src --threshold 0.8`: no duplicate candidates found.
- `mutator src/class_report/data.py --use-existing-coverage --test-command '.venv/bin/python -m pytest -q' --max-workers 2`: all 33 discovered mutations killed, none uncovered. Mutation testing was scoped to the data and statistics boundary, not the entire DOCX styling or CLI. These tools are locally available helpers and are not required to use the library or run CI.
- Tests include invalid scores, missing and duplicate headers, empty sheets and rows, missing games and categories, beta difficulty bounds, control characters, broken ZIP/XML input, Portuguese characters, statistics and DOCX values, deterministic paths, reexecution, replacement of output symlinks without modifying their targets, and installed CLI success/error behavior.

## Visual inspection

Generated `examples/alpha-report.docx` and `examples/beta-report.docx` from the synthetic XLSX fixtures with the bundled Python runtime (3.12.14). Rendered each with the documents runtime's `render_docx.py`, opened every resulting page image, removed an inherited title border, then regenerated, rendered, and inspected again. Each final example is one page. Confirmed readable accents, headings, table widths, cell padding, contrast, footer, and page number without visible clipping or overlap.

This verifies the supplied example layouts in the available renderer. Microsoft Word, assistive technology, physical printing, long documents, and every possible font substitution were not tested. Tables have repeated header rows and headings use Word paragraph styles, but this is not a formal accessibility certification. The project has no graphical frontend; the CLI and generated DOCX are its user interfaces.

## Reproducibility and limits

The runtime API only depends on python-docx and openpyxl. Rendering, artifact authoring, and local analysis tools are verification tooling and are not imported by the package. The development lock pins the environment used for verification. CI tests Python 3.11–3.13 on Linux and builds and smoke-tests the wheel; CI does not render documents.

Library generation expects records parsed by the loader for a single stage. The workbook is loaded read-only; formulas are not evaluated. Generation replaces individual report files atomically, not the whole batch. It is intended for trusted local course spreadsheets and dedicated output directories, not a multiuser upload service. The original arbitrary payment labels are retained; price categories are not forcibly reclassified.

## Repository and tutorial metadata

GitHub default branch verified as `master`. Updated and read back the English About description and these nine topics: `academic-reports`, `cli`, `docx`, `education`, `excel`, `game-development`, `openpyxl`, `python`, `python-docx`. README links only to the actual workflow and branch.

No Medium article or shortened tutorial URL was linked in the inspected original source/README, so no editorial publication was attempted. See `medium-update.pt-BR.md` for a concrete migration checklist if a corresponding article is subsequently identified. MIT license and the original author credit are preserved. No email, newsletter, package publication, or deployment was performed.
