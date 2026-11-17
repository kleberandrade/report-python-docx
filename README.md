# Class Report Generator

![Python 3.11+](https://img.shields.io/badge/Python-3.11%2B-3776AB)
![python-docx 1.2.0](https://img.shields.io/badge/python--docx-1.2.0-243746)
[![CI](https://github.com/kleberandrade/report-python-docx/actions/workflows/ci.yml/badge.svg?branch=master)](https://github.com/kleberandrade/report-python-docx/actions/workflows/ci.yml)
[![MIT](https://img.shields.io/badge/License-MIT-555555)](LICENSE)

Generate one Portuguese DOCX report per game from an Excel export of academic alpha or beta evaluations. The reusable Python package and CLI validate inputs before writing reports. Everything runs locally; there is no email delivery, file deletion command, plotting window, or external service.

## Requirements and installation

Python 3.11 or newer. Word, LibreOffice, a Google account, pandas, matplotlib, and wordcloud are **not** required to generate reports. Install from a checkout:

```sh
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-dev.lock
python -m pip install --no-build-isolation -e .
```

On Windows activate with `.venv\Scripts\activate`. The lock records the complete tested development environment. To install only runtime dependencies, use `python -m pip install .`. Runtime dependencies are python-docx 1.2.0 and openpyxl 3.1.5; dependency licenses remain their respective authors' licenses. No package has been published to PyPI by this project.

## Try the synthetic examples

```sh
class-report examples/alpha.xlsx --output reports/alpha
class-report examples/beta.xlsx --stage beta --output reports/beta
python -m class_report examples/alpha.xlsx --sheet Responses -o reports/alpha
```

The command reports the generated count and output directory. Invalid files, missing columns, invalid values, missing worksheets, and output permission failures exit with code 2 and an error message. It validates the whole worksheet before creating reports. Reports use deterministic sanitized names with a game-name hash, preventing path traversal and collisions between case variants. Running again replaces the corresponding files atomically. Unrelated files and reports for games removed from the input are retained. A batch is not transactional: a disk error may leave some completed reports; rerun after correcting the error. Use a dedicated output directory with one writer.

Open the included [alpha report](examples/alpha-report.docx) or [beta report](examples/beta-report.docx) for a preview. Both use fictional evaluations. Reports provide A4 margins, black semantic titles and headings, readable type, high-contrast tables with alternating rows and repeated headers, comments, and page numbers. Numeric tables replace the old charts and word clouds so values and original comments remain selectable. Empty comments display an explicit message. This is an academic evaluation layout, not a claim of ABNT compliance.

## Input contract

Use `.xlsx`, with headers on the first row of the active worksheet, or select a worksheet with `--sheet`. Header names are trimmed, case-sensitive Portuguese names. Additional unique columns are ignored. Duplicate headers are rejected. Entirely blank rows are ignored. Missing game names, scores, and categorical responses are errors. Comments may be blank. Scores must be finite numeric cells from 0 through 10; text numbers, booleans, formulas, and errors are rejected in numeric fields. Formulas are never evaluated. Save computed numeric results as values before importing.

Alpha requires:

- `Jogo avaliado`
- `Originalidade`, `Diversão`, `Mercado`
- `Você jogaria este jogo?` (`Sim` or `Não`)
- `Críticas e Sugestões` (may be blank)
- `Quanto pagaria por esta ideia?` (nonempty price label)

Beta requires:

- `Jogo avaliado`
- `Gráficos`, `Música e SFX`, `HUD`, `Jogabilidade`, `História`
- `Dificuldade` (numeric integer 1–5: Muito fácil, Fácil, Normal, Difícil, Muito difícil)
- `Ponto forte`, `Ponto fraco` (may be blank)
- `Idade` (finite number 0–120), `Sexo` (retained in the parsed record; not summarized in the report)

Scores use the mean and **population** standard deviation, preserving the original statistic. Percentages use all valid responses for that game as denominator and are rounded to one decimal, so totals may differ slightly from 100%. Comments preserve their order and deduplicate identical trimmed text. The input contains no evaluation date requirement; the report does not invent one. No final-stage schema is implemented because no corresponding source script or workbook was supplied.

## Library and architecture

```python
from pathlib import Path
from class_report import load_evaluations, generate_reports

items = load_evaluations(Path("examples/alpha.xlsx"), stage="alpha")
paths = generate_reports(items, Path("reports"))
```

`data.py` owns schemas, boundary validation, immutable evaluation records, and pure statistics. `report.py` turns validated records into styled DOCX documents and atomically replaces each output. `cli.py` owns argument parsing and user-facing errors. Library callers should obtain records through `load_evaluations`; `generate_reports` expects one homogeneous stage of validated records. A game is the exact trimmed `Jogo avaliado` string; respondent counts count rows, not unique people.

## Tests and build

```sh
python -m pytest --cov=class_report --cov-report=term-missing
ruff check src tests alpha_report.py beta_report_send_email.py
ruff format --check src tests alpha_report.py beta_report_send_email.py
python -m build --no-isolation
python -m pip install --force-reinstall --no-deps dist/*.whl
class-report examples/alpha.xlsx -o reports/wheel-smoke
```

CI runs these checks on Python 3.11, 3.12, and 3.13. See [verification](docs/verification.md) for actual local results, mutation scope, visual inspection, and limitations. The test suite uses synthetic workbooks, exercises the installed CLI, and checks DOCX content, invalid data, safe output paths, reruns, and symlink replacement. No Microsoft Word or physical device testing is claimed. Local dependency versions are pinned; DOCX ZIP metadata means byte-for-byte identical files are not guaranteed.

## Migration from the 2018 scripts

Install the package first. `python alpha_report.py` still uses `alpha evaluation.xlsx` by default; arguments are forwarded to the CLI. `python beta_report_send_email.py INPUT.xlsx -o reports/beta` keeps the old filename for compatibility but **only generates local beta reports**. The original beta workbook is not included. Importing either wrapper has no generation or delivery side effects. The old internal helper functions are replaced by the package API above.

The alpha interest calculation now correctly reads `Você jogaria este jogo?`; the old script mistakenly used the payment column. Output names now include a hash. There are no historical hardcoded survey dates, SMTP configuration, wildcard cleanup, or word clouds. Never put credentials in a workbook or committed source. Any formerly real embedded credentials should be rotated; this modernization does not rewrite Git history.

The original [alpha evaluation.xlsx](alpha%20evaluation.xlsx) is preserved for historical compatibility; it contains original responses and should be handled as course data, not copied into public examples. The modern synthetic examples contain no respondents or credentials.

## Presentation and teaching walkthrough

1. Open `examples/alpha.xlsx` and explain the original Portuguese Google Forms headers. Change a numeric score and save.
2. Run the alpha command above, then open the DOCX. Compare the criterion mean and population deviation with the two input rows.
3. Compare the independent interest and price tables; interest is 50% for each answer in the unchanged example.
4. Run the beta example and discuss difficulty codes, empty comments, and the full comments retained instead of word clouds.
5. Enter an invalid score or remove a required header and demonstrate the row-aware error. Restore the example afterward.
6. Run again to demonstrate stable filenames and replacement without deleting unrelated files.

Historical teaching images are retained in [Figures](Figures): [Excel export](Figures/Excel.PNG), [Spyder IDE](Figures/Spyder_IDE.PNG), and [old Word report](Figures/Word.PNG). Spyder and conda are optional; the commands above replace the unpinned 2018 installation instructions. `stopwords.txt` remains historical teaching material; the current generator does not need it. `clear.bat` is historical and must not be used for routine report generation because it deletes files in the current directory.

## Origin and license

Created by **Kleber de Oliveira Andrade**, 2018, for evaluations of games presented in the sixth-semester Digital Games for Console course at [Faculdade de Tecnologia de Americana](http://www.fatec.edu.br/). Original Portuguese project description: “Códigos para geração dos relatórios (avaliação alfa, beta e final) dos jogos apresentados na disciplina de Jogos Digitais para Console do curso de Jogos Digitais (6º semestre).”

MIT license and original copyright are preserved in [LICENSE](LICENSE). Historical screenshots and teaching assets retain their existing attribution. No Medium tutorial URL or shortener is linked by the inspected README or source files; see [tutorial review](docs/medium-update.pt-BR.md).
