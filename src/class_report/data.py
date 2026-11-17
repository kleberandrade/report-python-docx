"""Parse legacy Google Forms spreadsheets once, at the boundary."""

import csv
from collections import Counter
from dataclasses import dataclass
from math import isfinite
from pathlib import Path
from statistics import mean, pstdev

from openpyxl import load_workbook

GAME = "Jogo avaliado"
PLAY = "Você jogaria este jogo?"
PAY = "Quanto pagaria por esta ideia?"
DIFFICULTIES = ("Muito fácil", "Fácil", "Normal", "Difícil", "Muito difícil")


@dataclass(frozen=True)
class Schema:
    stage: str
    metrics: tuple[str, ...]
    categories: tuple[str, ...]
    comments: tuple[str, ...]
    demographics: tuple[str, ...] = ()

    @property
    def required(self):
        return (
            GAME,
            *self.metrics,
            *self.categories[:1],
            *self.comments,
            *self.categories[1:],
            *self.demographics,
        )


ALPHA = Schema(
    "alpha", ("Originalidade", "Diversão", "Mercado"), (PLAY, PAY), ("Críticas e Sugestões",)
)
BETA = Schema(
    "beta",
    ("Gráficos", "Música e SFX", "HUD", "Jogabilidade", "História"),
    ("Dificuldade",),
    ("Ponto forte", "Ponto fraco"),
    ("Idade", "Sexo"),
)
SCHEMAS = {"alpha": ALPHA, "beta": BETA}


@dataclass(frozen=True)
class Evaluation:
    game: str
    schema: Schema
    scores: tuple[float, ...]
    categories: tuple[str, ...]
    comments: tuple[str, ...]
    age: float | None = None
    gender: str = ""


@dataclass(frozen=True)
class Summary:
    means: tuple[float, ...]
    deviations: tuple[float, ...]
    play: dict[str, int]


def summarize(evaluations):
    columns = tuple(zip(*(item.scores for item in evaluations)))
    return Summary(
        tuple(map(mean, columns)),
        tuple(map(pstdev, columns)),
        dict(Counter(item.categories[0] for item in evaluations)),
    )


def text(value):
    result = str(value).strip() if value is not None else ""
    if any(ord(char) < 32 and char not in "\n\r\t" for char in result):
        raise ValueError("Unsupported control character")
    return result


def number(value, label, low, high):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{label}: expected a number from {low} to {high}")
    if not isfinite(value) or not low <= value <= high:
        raise ValueError(f"{label}: expected a number from {low} to {high}")
    return float(value)


def category(value, column):
    if column == "Dificuldade":
        index = number(value, column, 1, 5)
        if not index.is_integer():
            raise ValueError("Dificuldade: expected an integer from 1 to 5")
        return DIFFICULTIES[int(index) - 1]
    result = text(value)
    if not result:
        raise ValueError(f"{column}: required value")
    if column == PLAY and result not in ("Sim", "Não"):
        raise ValueError(f"{column}: expected Sim or Não")
    return result


def parse_row(values, schema):
    game = text(values[GAME])
    if not game:
        raise ValueError("Jogo avaliado: required value")
    age = None
    gender = ""
    if schema == BETA:
        age = number(values["Idade"], "Idade", 0, 120)
        gender = text(values["Sexo"])
    return Evaluation(
        game,
        schema,
        tuple(number(values[c], c, 0, 10) for c in schema.metrics),
        tuple(category(values[c], c) for c in schema.categories),
        tuple(text(values[c]) for c in schema.comments),
        age,
        gender,
    )


def read_rows(sheet, schema):
    rows = iter(sheet)
    headers = tuple(text(v) for v in next(rows, ()))
    if len(set(headers)) != len(headers):
        raise ValueError("Duplicate columns in worksheet")
    missing = set(schema.required) - set(headers)
    if missing:
        raise ValueError("Missing columns: " + ", ".join(sorted(missing)))
    result = []
    for index, row in enumerate(rows, 2):
        if all(v is None or v == "" for v in row):
            continue
        if len(row) != len(headers):
            raise ValueError(f"row {index}: expected {len(headers)} fields")
        try:
            result.append(parse_row(dict(zip(headers, row)), schema))
        except ValueError as error:
            raise ValueError(f"row {index}: {error}") from error
    if not result:
        raise ValueError("No evaluation rows found")
    return tuple(result)


def load_evaluations(path: Path, stage="alpha", sheet=None):
    schema = SCHEMAS[stage]
    suffix = Path(path).suffix.lower()
    if suffix == ".csv":
        if sheet is not None:
            raise ValueError("Worksheet selection is only available for XLSX input")
        return load_csv(Path(path), schema)
    if suffix != ".xlsx":
        raise ValueError("Input must be an .xlsx or .csv file")
    try:
        book = load_workbook(path, read_only=True, data_only=False, keep_links=False)
    except (KeyError, SyntaxError) as error:
        raise ValueError("Invalid XLSX workbook structure") from error
    try:
        if sheet is not None and sheet not in book.sheetnames:
            raise ValueError(f"Worksheet not found: {sheet}")
        return read_rows((book[sheet] if sheet else book.active).values, schema)
    finally:
        book.close()


def numeric_csv_rows(rows, schema):
    headers = next(rows, [])
    yield headers
    numeric = set(schema.metrics)
    if schema == BETA:
        numeric.update(("Dificuldade", "Idade"))
    positions = [i for i, header in enumerate(headers) if header.strip() in numeric]
    for row in rows:
        for index in positions:
            if index < len(row):
                try:
                    row[index] = float(row[index])
                except ValueError:
                    pass  # Boundary validation reports the row and criterion.
        yield row


def load_csv(path, schema):
    try:
        with path.open(encoding="utf-8-sig", newline="") as stream:
            return read_rows(numeric_csv_rows(csv.reader(stream, strict=True), schema), schema)
    except (csv.Error, UnicodeError) as error:
        raise ValueError("Invalid UTF-8 comma-separated CSV file") from error
