from __future__ import annotations

import argparse
import csv
import re
import sqlite3
import sys
from pathlib import Path


INTEGER_RE = re.compile(r"[+-]?\d+")


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--csv-dir", default="./output")
    parser.add_argument("--sqlite-path", default="./output/cbr_data.sqlite")
    args = parser.parse_args()

    csv_directory = Path(args.csv_dir)
    sqlite_path = Path(args.sqlite_path)

    if not csv_directory.is_dir():
        raise ValueError(f"CSV directory does not exist: {csv_directory}")

    csv_paths = sorted(csv_directory.glob("[0-9][0-9].*.csv"))
    if not csv_paths:
        raise ValueError(f"No CSV files found in {csv_directory}")

    table_names: set[str] = set()
    sqlite_path.parent.mkdir(parents=True, exist_ok=True)
    if sqlite_path.exists():
        sqlite_path.unlink()

    with sqlite3.connect(sqlite_path) as connection:
        for csv_path in csv_paths:
            table_name = build_table_name(csv_path)
            if table_name in table_names:
                raise ValueError(f"Duplicate table name generated: {table_name}")
            table_names.add(table_name)

            row_count = import_csv_to_table(connection, csv_path, table_name)
            print(f"{csv_path.name} -> {table_name} ({row_count} rows)")


def import_csv_to_table(connection: sqlite3.Connection, csv_path: Path, table_name: str) -> int:
    with csv_path.open("r", encoding="utf-8", newline="") as file:
        reader = csv.reader(file)
        rows = list(reader)

    if not rows:
        raise ValueError(f"CSV file is empty: {csv_path}")

    headers = rows[0]
    if not headers:
        raise ValueError(f"CSV file has no headers: {csv_path}")
    if len(set(headers)) != len(headers):
        raise ValueError(f"CSV file contains duplicate headers: {csv_path}")

    data_rows = rows[1:]
    column_types = infer_column_types(headers, data_rows)

    create_table_sql = build_create_table_sql(table_name, headers, column_types)
    connection.execute(create_table_sql)

    placeholders = ", ".join("?" for _ in headers)
    insert_sql = f"INSERT INTO {quote_identifier(table_name)} VALUES ({placeholders})"
    converted_rows = [convert_row(row, column_types) for row in data_rows]
    connection.executemany(insert_sql, converted_rows)
    connection.commit()

    return len(data_rows)


def infer_column_types(headers: list[str], data_rows: list[list[str]]) -> list[str]:
    column_types: list[str] = []

    for column_index, _header in enumerate(headers):
        values = [row[column_index] for row in data_rows if column_index < len(row) and row[column_index] != ""]
        column_types.append(infer_column_type(values))

    return column_types


def infer_column_type(values: list[str]) -> str:
    if not values:
        return "TEXT"

    if any(is_integer_with_leading_zero(value) for value in values):
        return "TEXT"

    if all(is_integer(value) for value in values):
        return "INTEGER"

    if all(is_real(value) for value in values):
        return "REAL"

    return "TEXT"


def is_integer(value: str) -> bool:
    return bool(INTEGER_RE.fullmatch(value))


def is_integer_with_leading_zero(value: str) -> bool:
    normalized = value.lstrip("+-")
    return len(normalized) > 1 and normalized.startswith("0") and normalized.isdigit()


def is_real(value: str) -> bool:
    try:
        float(value)
    except ValueError:
        return False

    return True


def convert_row(row: list[str], column_types: list[str]) -> list[object | None]:
    if len(row) != len(column_types):
        raise ValueError("CSV row length does not match header length.")

    converted: list[object | None] = []
    for value, column_type in zip(row, column_types, strict=True):
        if value == "":
            converted.append(None)
        elif column_type == "INTEGER":
            converted.append(int(value))
        elif column_type == "REAL":
            converted.append(float(value))
        else:
            converted.append(value)

    return converted


def build_create_table_sql(table_name: str, headers: list[str], column_types: list[str]) -> str:
    column_definitions = ", ".join(
        f"{quote_identifier(header)} {column_type}"
        for header, column_type in zip(headers, column_types, strict=True)
    )
    return f"CREATE TABLE {quote_identifier(table_name)} ({column_definitions})"


def build_table_name(csv_path: Path) -> str:
    stem_without_numeric_prefix = re.sub(r"^\d+\.", "", csv_path.stem)
    normalized_stem = "".join(
        character.lower() if character.isalnum() else "_"
        for character in stem_without_numeric_prefix
    )
    normalized_stem = re.sub(r"_+", "_", normalized_stem).strip("_")
    if not normalized_stem:
        raise ValueError(f"Cannot build table name from {csv_path}")
    return normalized_stem


def quote_identifier(identifier: str) -> str:
    return '"' + identifier.replace('"', '""') + '"'


if __name__ == "__main__":
    main()
