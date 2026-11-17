import subprocess
import sys

import pytest
from docx import Document
from openpyxl import Workbook

from class_report.data import ALPHA, BETA, load_evaluations, summarize
from class_report.report import generate_reports


def workbook(tmp_path, headers=None, rows=None):
    book = Workbook()
    book.active.append(headers or list(ALPHA.required))
    for row in rows or [
        ["Órbita", 8, 6, 10, "Sim", "Ótima ação", "Menos R$ 10,00"],
        ["Órbita", 6, 10, 8, "Não", None, "Menos R$ 10,00"],
    ]:
        book.active.append(row)
    path = tmp_path / "input.xlsx"
    book.save(path)
    return path


def test_alpha_statistics_and_document(tmp_path):
    evaluations = load_evaluations(workbook(tmp_path), "alpha")
    summary = summarize(evaluations)
    assert summary.means == (7, 8, 9)
    assert summary.deviations == (1, 2, 1)
    assert summary.play == {"Sim": 1, "Não": 1}
    paths = generate_reports(evaluations, tmp_path / "out")
    doc = Document(paths[0])
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Órbita" in text and "Ótima ação" in text
    assert "None" not in text
    assert doc.tables[1].cell(1, 2).text == "50.0%"
    assert doc.tables[0].cell(1, 1).text == "7.00"
    assert generate_reports(evaluations, tmp_path / "out") == paths
    assert len(list((tmp_path / "out").glob("*.docx"))) == 1


@pytest.mark.parametrize("bad", [None, -1, 11, True, "8", "=1+2", float("inf")])
def test_invalid_score(tmp_path, bad):
    with pytest.raises(ValueError, match="row 2.*Originalidade"):
        load_evaluations(
            workbook(tmp_path, rows=[["Jogo", bad, 8, 8, "Sim", "", "Preço"]]), "alpha"
        )


def test_missing_headers_and_empty_data(tmp_path):
    with pytest.raises(ValueError, match="columns"):
        load_evaluations(workbook(tmp_path, headers=["Game"]), "alpha")
    path = workbook(tmp_path)
    book = Workbook()
    book.active.append(list(ALPHA.required))
    book.save(path)
    with pytest.raises(ValueError, match="No evaluation"):
        load_evaluations(path, "alpha")


def test_beta(tmp_path):
    rows = [["Beta", 8, 7, 9, 10, 6, 3, "Boa música", None, 20, "Outro"]]
    items = load_evaluations(workbook(tmp_path, list(BETA.required), rows), "beta")
    doc = Document(generate_reports(items, tmp_path / "beta")[0])
    text = "\n".join(p.text for p in doc.paragraphs)
    assert "Boa música" in text and "Nenhum comentário" in text
    assert "Normal" in "\n".join(c.text for t in doc.tables for r in t.rows for c in r.cells)


def test_filename_does_not_escape_and_games_do_not_collide(tmp_path):
    rows = [[name, 8, 8, 8, "Sim", "", "Preço"] for name in ["../a", "..\\a", "A", "a"]]
    paths = generate_reports(
        load_evaluations(workbook(tmp_path, rows=rows), "alpha"), tmp_path / "out"
    )
    assert len(set(paths)) == 4
    assert all(p.parent == tmp_path / "out" for p in paths)


def test_cli_errors_and_success(tmp_path):
    command = [sys.executable, "-m", "class_report"]
    result = subprocess.run(
        command + [str(tmp_path / "missing.xlsx")], capture_output=True, text=True
    )
    assert result.returncode == 2 and "error:" in result.stderr and "Traceback" not in result.stderr
    result = subprocess.run(
        command + [str(workbook(tmp_path)), "-o", str(tmp_path / "reports")],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0 and "1 report" in result.stdout
    assert len(list((tmp_path / "reports").glob("*.docx"))) == 1


@pytest.mark.parametrize("column,value", [(0, ""), (4, "Talvez"), (6, None)])
def test_required_categories(tmp_path, column, value):
    row = ["Jogo", 0, 10, 5, "Sim", "", "Preço"]
    row[column] = value
    with pytest.raises(ValueError):
        load_evaluations(workbook(tmp_path, rows=[row]))


def test_duplicate_headers_sheet_selection_and_suffix(tmp_path):
    with pytest.raises(ValueError, match="Duplicate"):
        load_evaluations(workbook(tmp_path, headers=["A", "A"]))
    with pytest.raises(ValueError, match="Worksheet not found"):
        load_evaluations(workbook(tmp_path), sheet="Missing")
    assert len(load_evaluations(workbook(tmp_path), sheet="Sheet")) == 2
    with pytest.raises(ValueError, match=".xlsx"):
        load_evaluations(tmp_path / "a.csv")
    with pytest.raises(ValueError, match="No evaluation"):
        generate_reports([], tmp_path / "out")


@pytest.mark.parametrize("difficulty", [0, 6, 2.5])
def test_invalid_beta_difficulty(tmp_path, difficulty):
    row = ["Beta", 8, 7, 9, 10, 6, difficulty, "", "", 20, "Outro"]
    with pytest.raises(ValueError, match="Dificuldade"):
        load_evaluations(workbook(tmp_path, list(BETA.required), [row]), "beta")


def test_blank_rows_and_score_endpoints(tmp_path):
    rows = [[None] * 7, ["Jogo", 0, 10, 5, "Sim", "", "Preço"]]
    items = load_evaluations(workbook(tmp_path, rows=rows))
    assert items[0].scores == (0, 10, 5) and len(items) == 1


def test_cli_boundary_direct(tmp_path, capsys):
    from class_report.cli import main

    assert main([str(workbook(tmp_path)), "-o", str(tmp_path / "cli")]) == 0
    assert "1 report" in capsys.readouterr().out
    with pytest.raises(SystemExit) as error:
        main([str(tmp_path / "missing.xlsx")])
    assert error.value.code == 2


def test_broken_workbook_has_friendly_error(tmp_path, capsys):
    from class_report.cli import main

    path = tmp_path / "broken.xlsx"
    path.write_text("not a zip")
    with pytest.raises(SystemExit) as error:
        main([str(path)])
    assert error.value.code == 2 and "error:" in capsys.readouterr().err


def test_reports_do_not_follow_existing_symlink(tmp_path):
    victim = tmp_path / "private.txt"
    victim.write_text("unchanged")
    items = load_evaluations(workbook(tmp_path))
    output = tmp_path / "out"
    paths = generate_reports(items, output)
    paths[0].unlink()
    paths[0].symlink_to(victim)
    generate_reports(items, output)
    assert victim.read_text() == "unchanged"
    assert not paths[0].is_symlink()


def test_missing_columns_without_duplicate_headers(tmp_path):
    with pytest.raises(ValueError, match="Missing columns:"):
        load_evaluations(workbook(tmp_path, headers=["Game"], rows=[["Name"]]))


def test_control_characters_rejected():
    from class_report.data import text

    with pytest.raises(ValueError, match="control character"):
        text("bad\x00value")
    assert text("ação\nportuguesa") == "ação\nportuguesa"


def test_title_has_no_decorative_border(tmp_path):
    from docx.oxml.ns import qn

    doc = Document(generate_reports(load_evaluations(workbook(tmp_path)), tmp_path / "out")[0])
    assert not list(doc.styles["Title"].element.iter(qn("w:pBdr")))


def test_invalid_ooxml_has_friendly_error(tmp_path, capsys):
    from zipfile import ZipFile
    from class_report.cli import main

    path = tmp_path / "invalid.xlsx"
    with ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", "<broken")
    with pytest.raises(SystemExit) as error:
        main([str(path)])
    assert error.value.code == 2 and "error:" in capsys.readouterr().err
