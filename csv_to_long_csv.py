from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path

from cbr_data_parser.utils.long_subject_tables import (
    LONG_SUBJECT_TABLES,
    REGIONAL_DIMENSION_HEADERS,
    classify_long_subject_table,
    resolve_regional_dimensions,
    resolve_section_type,
)


COMMON_PREFIX_HEADER = "Раздел"
COMMON_MIDDLE_HEADERS = ["Тип раздела"]
INDICATOR_LEVEL_HEADERS = [
    "Показатель 1",
    "Показатель 2",
    "Показатель 3",
    "Показатель 4",
    "Показатель 5",
]
COMMON_SUFFIX_HEADERS = ["Значение", "ед. Измерения"]


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--csv-dir", default="./output")
    parser.add_argument("--output-dir", default="./output_long")
    args = parser.parse_args()

    csv_directory = Path(args.csv_dir)
    output_directory = Path(args.output_dir)

    if not csv_directory.is_dir():
        raise ValueError(f"CSV directory does not exist: {csv_directory}")

    csv_paths = sorted(csv_directory.glob("[0-9][0-9].*.csv"))
    if not csv_paths:
        raise ValueError(f"No CSV files found in {csv_directory}")

    output_directory.mkdir(parents=True, exist_ok=True)
    remove_existing_long_csvs(output_directory)

    grouped_headers: dict[str, list[str]] = {}
    grouped_rows: dict[str, list[list[str]]] = {
        spec["key"]: []
        for spec in LONG_SUBJECT_TABLES
    }

    for csv_path in csv_paths:
        metadata_path = csv_path.with_suffix(".json")
        if not metadata_path.is_file():
            raise ValueError(f"Metadata JSON does not exist: {metadata_path}")

        subject_key, output_headers, output_rows = convert_csv_to_long_rows(
            csv_path=csv_path,
            metadata_path=metadata_path,
        )

        if subject_key in grouped_headers:
            if grouped_headers[subject_key] != output_headers:
                raise ValueError(
                    f'Inconsistent headers for subject table "{subject_key}": '
                    f"{grouped_headers[subject_key]} != {output_headers}"
                )
        else:
            grouped_headers[subject_key] = output_headers

        grouped_rows[subject_key].extend(output_rows)
        print(f"{csv_path.name} -> {subject_key} ({len(output_rows)} rows)")

    for spec in LONG_SUBJECT_TABLES:
        subject_key = spec["key"]
        output_rows = grouped_rows[subject_key]
        if not output_rows:
            continue

        output_path = output_directory / spec["file_name"]
        write_long_csv(
            output_path=output_path,
            output_headers=grouped_headers[subject_key],
            output_rows=output_rows,
        )
        print(f"{spec['file_name']} ({len(output_rows)} rows)")


def remove_existing_long_csvs(output_directory: Path) -> None:
    for csv_path in output_directory.glob("[0-9][0-9].*.csv"):
        csv_path.unlink()


def convert_csv_to_long_rows(
    csv_path: Path,
    metadata_path: Path,
) -> tuple[str, list[str], list[list[str]]]:
    metadata = json.loads(metadata_path.read_text(encoding="utf-8"))

    with csv_path.open("r", encoding="utf-8", newline="") as file:
        reader = csv.reader(file)
        rows = list(reader)

    if not rows:
        raise ValueError(f"CSV file is empty: {csv_path}")

    headers = rows[0]
    data_rows = rows[1:]

    expected_prefix_headers = get_expected_prefix_headers(metadata)
    actual_prefix_headers = headers[: len(expected_prefix_headers)]
    if actual_prefix_headers != expected_prefix_headers:
        raise ValueError(
            f"CSV headers do not match metadata in {csv_path}. "
            f"Expected prefix {expected_prefix_headers}, got {actual_prefix_headers}."
        )

    value_headers = headers[len(expected_prefix_headers) :]
    if not value_headers:
        raise ValueError(f"CSV file does not contain value columns: {csv_path}")

    subject_key = classify_long_subject_table(metadata)
    output_headers = build_output_headers(metadata)
    output_rows: list[list[str]] = []
    regional_hierarchy_state = {"country": None, "district": None}
    indicator_paths = build_indicator_paths(metadata)

    for row in data_rows:
        if len(row) != len(headers):
            raise ValueError(f"CSV row length does not match header length in {csv_path}")

        section, dimension_values = build_output_prefix_values(
            metadata=metadata,
            row=row,
            regional_hierarchy_state=regional_hierarchy_state,
        )
        section_type = resolve_section_type(section)
        for value_header, value in zip(value_headers, row[len(expected_prefix_headers) :], strict=True):
            if value == "":
                continue

            indicator_levels = normalize_indicator_levels(indicator_paths[value_header])
            unit = resolve_unit(
                metadata=metadata,
                section=section,
                value_header=value_header,
            )
            output_rows.append([section, section_type, *dimension_values, *indicator_levels, value, unit])

    return subject_key, output_headers, output_rows


def write_long_csv(output_path: Path, output_headers: list[str], output_rows: list[list[str]]) -> None:
    with output_path.open("w", encoding="utf-8", newline="") as file:
        writer = csv.writer(file)
        writer.writerow(output_headers)
        writer.writerows(output_rows)


def build_output_headers(metadata: dict) -> list[str]:
    headers = [COMMON_PREFIX_HEADER]
    headers.extend(COMMON_MIDDLE_HEADERS)

    if not metadata.get("row_id_column"):
        headers.extend(get_expected_prefix_headers(metadata))
        if metadata.get("row_label_column"):
            headers.extend(REGIONAL_DIMENSION_HEADERS)

    headers.extend(INDICATOR_LEVEL_HEADERS)
    headers.extend(COMMON_SUFFIX_HEADERS)
    return headers


def get_expected_prefix_headers(metadata: dict) -> list[str]:
    row_id_column = metadata.get("row_id_column")
    if row_id_column:
        return [row_id_column]

    row_label_column = metadata.get("row_label_column")
    if row_label_column:
        return [row_label_column["label"]]

    id_columns = metadata.get("id_columns", [])
    if id_columns:
        return [column["label"] for column in id_columns]

    raise ValueError("Metadata does not contain row identifier columns.")


def build_output_prefix_values(
    metadata: dict,
    row: list[str],
    regional_hierarchy_state: dict[str, str | None],
) -> tuple[str, list[str]]:
    row_id_column = metadata.get("row_id_column")
    if row_id_column:
        return row[0], []

    row_label_column = metadata.get("row_label_column")
    if row_label_column:
        return build_metric_label(metadata), [
            row[0],
            *resolve_regional_dimensions(row[0], regional_hierarchy_state),
        ]

    id_columns = metadata.get("id_columns", [])
    if id_columns:
        return build_metric_label(metadata), row[: len(id_columns)]

    raise ValueError("Metadata does not contain row identifier columns.")


def build_metric_label(metadata: dict) -> str:
    description = str(metadata["sheet_description"]).strip().rstrip(".")

    for suffix in [", тыс руб", ", тыс руб.", ", ед", ", ед.", ", единиц"]:
        if description.endswith(suffix):
            return description[: -len(suffix)]

    return description


def build_indicator_paths(metadata: dict) -> dict[str, list[str]]:
    indicator_columns = metadata.get("insurance_columns") or metadata.get("metric_columns")
    if not indicator_columns:
        raise ValueError("Metadata does not contain indicator columns.")

    indicator_paths: dict[str, list[str]] = {}
    for column in indicator_columns:
        column_name = column["column_name"]
        if column_name in indicator_paths:
            raise ValueError(f'Duplicate indicator column "{column_name}" in metadata.')
        indicator_paths[column_name] = list(column["path"])

    return indicator_paths


def normalize_indicator_levels(path: list[str]) -> list[str]:
    if len(path) > len(INDICATOR_LEVEL_HEADERS):
        raise ValueError(
            f"Indicator path depth {len(path)} exceeds supported depth {len(INDICATOR_LEVEL_HEADERS)}."
        )

    return [*([""] * (len(INDICATOR_LEVEL_HEADERS) - len(path))), *path]


def resolve_unit(metadata: dict, section: str, value_header: str) -> str:
    row_id_column = metadata.get("row_id_column")
    row_label_column = metadata.get("row_label_column")
    insurance_columns = metadata.get("insurance_columns")
    metric_columns = metadata.get("metric_columns")

    if row_id_column:
        unit = extract_unit(section)
        if unit is None:
            raise ValueError(f'Cannot determine unit from summary row label "{section}"')
        return unit

    if row_label_column:
        unit = extract_unit(str(metadata["sheet_description"]))
        if unit is None:
            raise ValueError(f'Cannot determine unit from sheet description "{metadata["sheet_description"]}"')
        return unit

    if insurance_columns is not None:
        unit = extract_unit(str(metadata["sheet_description"]))
        if unit is None:
            raise ValueError(f'Cannot determine unit from sheet description "{metadata["sheet_description"]}"')
        return unit

    if metric_columns is not None:
        unit = extract_unit(value_header)
        if unit is not None:
            return unit

        unit = extract_unit(str(metadata["sheet_description"]))
        if unit is not None:
            return unit

        raise ValueError(
            f'Cannot determine unit from metric header "{value_header}" '
            f'or sheet description "{metadata["sheet_description"]}"'
        )

    raise ValueError("Metadata does not contain recognized value columns.")


def extract_unit(text: str) -> str | None:
    normalized = text.lower().replace(",", " ").replace("\xa0", " ")

    if "%" in normalized:
        return "%"
    if "тыс" in normalized and "руб" in normalized:
        return "тыс. руб."
    if "чел" in normalized:
        return "чел."
    if "единиц" in normalized or " ед." in normalized or normalized.endswith(" ед") or normalized.endswith(" ед."):
        return "ед."
    if normalized.endswith("ед"):
        return "ед."

    return None


if __name__ == "__main__":
    main()
