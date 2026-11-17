"""Readable academic DOCX reports; no network or email side effects."""

from collections import Counter
from hashlib import sha256
from pathlib import Path
import re
import tempfile

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .data import summarize


def configure(document):
    section = document.sections[0]
    section.page_width, section.page_height = Cm(21), Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(2)
    section.left_margin = section.right_margin = Cm(2.2)
    for name in ("Normal", "Title", "Subtitle", "Heading 1", "Heading 2"):
        style = document.styles[name]
        style.font.name = "Calibri"
        style.font.color.rgb = RGBColor(0, 0, 0)
    normal = document.styles["Normal"]
    normal.font.size = Pt(11)
    normal.paragraph_format.space_after = Pt(7)
    normal.paragraph_format.line_spacing = 1.1
    document.styles["Title"].font.size = Pt(26)
    for border in list(document.styles["Title"].element.iter(qn("w:pBdr"))):
        border.getparent().remove(border)
    document.styles["Heading 1"].font.size = Pt(15)
    section.header.paragraphs[0].text = "AVALIAÇÃO ACADÊMICA  /  JOGOS DIGITAIS"
    footer = section.footer.paragraphs[0]
    footer.text = "Relatório de avaliação • Página "
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    document.core_properties.author = "Kleber de Oliveira Andrade"
    document.core_properties.language = "pt-BR"


def style_cell(cell, fill, color):
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    properties = cell._tc.get_or_add_tcPr()
    shade = OxmlElement("w:shd")
    shade.set(qn("w:fill"), fill)
    properties.append(shade)
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        edge = OxmlElement("w:" + side)
        for key, value in (("val", "single"), ("sz", "4"), ("color", "D9D9D9")):
            edge.set(qn("w:" + key), value)
        borders.append(edge)
    properties.append(borders)
    margins = OxmlElement("w:tcMar")
    for side in ("top", "left", "bottom", "right"):
        edge = OxmlElement("w:" + side)
        edge.set(qn("w:w"), "100")
        edge.set(qn("w:type"), "dxa")
        margins.append(edge)
    properties.append(margins)
    for paragraph in cell.paragraphs:
        paragraph.paragraph_format.space_after = Pt(3)
        paragraph.paragraph_format.space_before = Pt(3)
        for run in paragraph.runs:
            run.font.color.rgb = RGBColor.from_string(color)


def table(document, headers, rows, widths):
    result = document.add_table(rows=1, cols=len(headers))
    result.alignment = WD_TABLE_ALIGNMENT.CENTER
    result.autofit = False
    for column, width in zip(result.columns, widths):
        column.width = Cm(width)
    for cell, label in zip(result.rows[0].cells, headers):
        cell.text = label
        style_cell(cell, "243746", "FFFFFF")
    repeat = OxmlElement("w:tblHeader")
    result.rows[0]._tr.get_or_add_trPr().append(repeat)
    for index, values in enumerate(rows):
        row = result.add_row()
        for cell, value, width in zip(row.cells, values, widths):
            cell.width = Cm(width)
            cell.text = str(value)
            style_cell(cell, "F1F4F6" if index % 2 == 0 else "FFFFFF", "000000")
    document.add_paragraph()
    return result


def build_document(items):
    first = items[0]
    document = Document()
    configure(document)
    document.add_paragraph("Relatório de avaliação do jogo", "Title")
    document.add_paragraph(first.game, "Subtitle")
    document.add_paragraph(f"Versão {first.schema.stage} • {len(items)} avaliações")
    document.add_paragraph(
        "Este relatório reúne as avaliações dos participantes para apoiar "
        "a revisão do jogo. As notas variam de 0 a 10; valores maiores "
        "indicam avaliações mais favoráveis."
    )
    stats = summarize(items)
    document.add_heading("Avaliação por critério", 1)
    table(
        document,
        ("Critério", "Média", "Desvio padrão"),
        [
            (name, f"{avg:.2f}", f"{std:.2f}")
            for name, avg, std in zip(first.schema.metrics, stats.means, stats.deviations)
        ],
        (8, 3, 5.6),
    )
    document.add_paragraph(
        "Desvio padrão populacional calculado sobre as respostas deste jogo. "
        "As médias não representam uma amostra de toda a população."
    )
    add_categories(document, items)
    add_comments(document, items)
    return document


def add_categories(document, items):
    for index, column in enumerate(items[0].schema.categories):
        document.add_heading(column.replace("?", ""), 1)
        counts = Counter(item.categories[index] for item in items)
        table(
            document,
            ("Resposta", "Pessoas", "Percentual"),
            [
                (name, count, f"{count / len(items) * 100:.1f}%")
                for name, count in sorted(counts.items())
            ],
            (10, 2.8, 3.8),
        )
    if items[0].age is not None:
        average = sum(item.age for item in items) / len(items)
        document.add_paragraph(f"Idade média dos participantes: {average:.1f} anos.")


def add_comments(document, items):
    for index, heading in enumerate(items[0].schema.comments):
        document.add_heading(heading, 1)
        comments = dict.fromkeys(item.comments[index] for item in items if item.comments[index])
        if not comments:
            document.add_paragraph("Nenhum comentário informado.")
        for comment in comments:
            document.add_paragraph(comment, "List Bullet")


def report_name(game):
    slug = re.sub(r"[^\w-]+", "-", game, flags=re.UNICODE).strip("-_")[:70] or "game"
    return f"{slug}-{sha256(game.encode()).hexdigest()[:12]}.docx"


def save_atomic(document, target):
    with tempfile.NamedTemporaryFile(dir=target.parent, suffix=".docx", delete=False) as handle:
        temporary = Path(handle.name)
    try:
        document.save(temporary)
        temporary.replace(target)
    finally:
        temporary.unlink(missing_ok=True)


def generate_reports(evaluations, output):
    if not evaluations:
        raise ValueError("No evaluation rows found")
    output = Path(output)
    output.mkdir(parents=True, exist_ok=True)
    groups = {}
    for item in evaluations:
        groups.setdefault(item.game, []).append(item)
    paths = []
    for game, items in sorted(groups.items()):
        target = output / report_name(game)
        save_atomic(build_document(items), target)
        paths.append(target)
    return paths
