from __future__ import annotations

import argparse
import json
import re
import sqlite3
import sys
from pathlib import Path

from cbr_data_parser.utils.long_subject_tables import (
    LONG_SUBJECT_TABLES,
    SECTION_TYPE_DESCRIPTIONS,
    classify_long_subject_table,
)


TABLE_DESCRIPTIONS = {
    "summary": (
        "This table combines the two summary sheets that do not have a separate subject dimension. "
        "It covers the overall activity KPIs and the summary by key insurance lines."
    ),
    "insurers": (
        "This table combines all insurer-level sheets. It includes core insurance metrics, "
        "inward and outward reinsurance, ОСАГО, direct reimbursement, ОМС, and intermediary participation."
    ),
    "regions": (
        "This table combines all regional sheets for federal districts and subjects of the Russian Federation. "
        "It includes premiums, contract counts, contract amounts, payouts, and settled claims."
    ),
    "mutual_societies": (
        "This table contains the OVS subject table. It includes membership counts, entries, exits, "
        "closing balances, and related membership indicators."
    ),
}

ROW_GRAINS = {
    "summary": "one summary-section and indicator pair per row",
    "insurers": "one insurer, section, and indicator pair per row",
    "regions": "one region, section, and indicator pair per row",
    "mutual_societies": "one OVS entity, section, and indicator pair per row",
}

INDICATOR_DESCRIPTIONS = {
    "summary": (
        "contains insurance types, summary KPI rows, and hierarchical insurance categories from the two transposed summary sheets"
    ),
    "insurers": (
        "contains all insurer-level indicators from the source workbook, including insurance-line hierarchies, "
        "reinsurance metrics, ОСАГО and ОМС KPIs, and intermediary-related measures"
    ),
    "regions": (
        "contains regional indicators for insurance lines and comparable regional KPI blocks across all regional sheets"
    ),
    "mutual_societies": (
        "contains OVS membership indicators, including opening balance, joins, exits by reason, closing balance, and related counts"
    ),
}


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")

    parser = argparse.ArgumentParser()
    parser.add_argument("--metadata-dir", default="./output")
    parser.add_argument("--sqlite-path", default="./output/cbr_data.sqlite")
    parser.add_argument("--output-path", default="./bi_prompt.txt")
    args = parser.parse_args()

    prompt = build_bi_sql_prompt(
        metadata_dir=Path(args.metadata_dir),
        sqlite_path=Path(args.sqlite_path),
    )

    output_path = Path(args.output_path)
    output_path.write_text(prompt, encoding="utf-8")
    print(prompt)


def build_bi_sql_prompt(metadata_dir: Path, sqlite_path: Path) -> str:
    table_specs = load_table_specs(metadata_dir=metadata_dir, sqlite_path=sqlite_path)
    section_type_level_counts = load_section_type_level_counts(sqlite_path=sqlite_path)

    lines: list[str] = [
        f'You work with the SQLite database at "{sqlite_path.resolve()}".',
        "",
        "The database uses a long analytical structure. Table descriptions below are based on the workbook metadata and the actual SQLite schemas.",
        "",
        "Core SQL rules:",
        '- Always quote table names and column names with double quotes.',
        '- All analytical tables are narrow and row-based. The main measure value is in `"Значение"`.',
        '- The detailed analytical dimension is split across `"Показатель 1"` ... `"Показатель 5"`.',
        '- Indicator paths can have different depth.',
        '- Indicator levels are right-aligned: the deepest level is always stored in the last non-empty `"Показатель N"` column, and shorter paths leave the leading level columns empty.',
        '- Use only the non-empty level columns when matching an indicator path.',
        '- `"Тип раздела"` stores a normalized business category for `"Раздел"` such as `"страховая премия"`, `"размер выплаты"`, or `"количество договоров"`.',
        '- `"ед. Измерения"` stores the unit and should be used when mixing or comparing indicators.',
        '- `"Раздел"` stores the source metric block. In `"сводные_показатели"` it is the original KPI row. In the subject tables it is the source sheet description without the trailing unit suffix.',
        '- In `"регионы"`, the raw regional label is kept in `"Наименования федеральных округов и субъектов Российской Федерации"`, and the hierarchy is split into `"country"`, `"district"`, and `"region"`.',
        '- Treat NULL as missing data, not as zero.',
        '- Prefer filtering on `"Тип раздела"` when you need a broad metric family and on `"Раздел"` when you need the exact source block.',
        '- Do not aggregate different `"ед. Измерения"` values together.',
        '- Do not aggregate different `"Тип раздела"`, `"Раздел"`, or indicator level combinations together unless you explicitly intend to roll them up.',
        "",
        "Table catalog:",
    ]

    for spec in table_specs:
        lines.append(
            f'- "{spec["table_name"]}": {spec["description"]} '
            f'Row grain: {spec["row_grain"]}. '
            f'Identifier columns: {spec["identifier_columns"]}. '
            f'Fields `"Показатель 1"` ... `"Показатель 5"`: {spec["indicator_description"]}. '
            f'Actual columns: {spec["actual_columns"]}.'
        )

    lines.append("")
    lines.append('Available `"Тип раздела"` values:')
    for section_type, description in SECTION_TYPE_DESCRIPTIONS.items():
        level_counts = section_type_level_counts.get(section_type)
        if level_counts:
            level_counts_text = format_level_counts(level_counts)
            lines.append(
                f'- "{section_type}": {description} Possible indicator levels: {level_counts_text}.'
            )
        else:
            lines.append(f'- "{section_type}": {description}')

    lines.extend(
        [
            "",
            "Query rules:",
            '- Most comparisons now happen inside one subject table by filtering on `"Раздел"`.',
            '- Use `"Тип раздела"` for broad groupings such as premiums, payouts, contract counts, or insured-person counts.',
            '- In `"страховщики"`, identify entities by `"Регистрационный номер записи страховщика в едином государственном реестре субъектов страхового дела"` and keep `"Полное наименование страховщика"` as a label.',
            '- In `"регионы"`, use `"country"` for the top territorial split, `"district"` for federal districts, and `"region"` for specific subjects; the original row label remains in `"Наименования федеральных округов и субъектов Российской Федерации"`.',
            '- In `"общества_взаимного_страхования"`, identify rows by `"Регистрационный номер записи общества взаимного страхования в едином государственном реестре субъектов страхового дела"`.',
            '- When you compare two different metric blocks inside the same table, match on the subject identifier columns and, when needed, on all indicator level columns.',
            '- When you combine results, also keep `"ед. Измерения"` aligned.',
            '- Do not compare `"сводные_показатели"` to the subject tables unless you explicitly map the business meaning of both `"Раздел"` and the indicator level path.',
            "",
            "Example query patterns:",
            "",
            "1. Top insurers for one chosen metric block and indicator:",
            "```sql",
            "SELECT",
            '  "Регистрационный номер записи страховщика в едином государственном реестре субъектов страхового дела" AS insurer_id,',
            '  "Полное наименование страховщика" AS insurer_name,',
            '  "Тип раздела",',
            '  "Раздел",',
            '  "Показатель 1",',
            '  "Показатель 2",',
            '  "Показатель 3",',
            '  "Показатель 4",',
            '  "Показатель 5",',
            '  "Значение",',
            '  "ед. Измерения"',
            'FROM "страховщики"',
            'WHERE "Тип раздела" = \'страховая премия\'',
            '  AND "Раздел" = \'Сведения о страховых премиях (взносах) по договорам страхования, в разрезе страховщиков\'',
            '  AND "Показатель 5" = \'Всего*\'',
            '  AND "Показатель 1" IS NULL',
            '  AND "Показатель 2" IS NULL',
            '  AND "Показатель 3" IS NULL',
            '  AND "Показатель 4" IS NULL',
            'ORDER BY "Значение" DESC',
            "LIMIT 20;",
            "```",
            "",
            "2. Compare two insurer metric blocks inside the same table:",
            "```sql",
            "SELECT",
            '  p."Регистрационный номер записи страховщика в едином государственном реестре субъектов страхового дела" AS insurer_id,',
            '  p."Полное наименование страховщика" AS insurer_name,',
            '  p."Тип раздела" AS premium_section_type,',
            '  p."Показатель 1",',
            '  p."Показатель 2",',
            '  p."Показатель 3",',
            '  p."Показатель 4",',
            '  p."Показатель 5",',
            '  p."Значение" AS premium_value,',
            '  q."Значение" AS payout_value,',
            "  CASE",
            '    WHEN p."Значение" IS NULL OR p."Значение" = 0 THEN NULL',
            '    ELSE 1.0 * q."Значение" / p."Значение"',
            "  END AS payout_ratio",
            'FROM "страховщики" p',
            'JOIN "страховщики" q',
            '  ON p."Регистрационный номер записи страховщика в едином государственном реестре субъектов страхового дела" =',
            '     q."Регистрационный номер записи страховщика в едином государственном реестре субъектов страхового дела"',
            ' AND p."Показатель 1" IS q."Показатель 1"',
            ' AND p."Показатель 2" IS q."Показатель 2"',
            ' AND p."Показатель 3" IS q."Показатель 3"',
            ' AND p."Показатель 4" IS q."Показатель 4"',
            ' AND p."Показатель 5" IS q."Показатель 5"',
            'WHERE p."Тип раздела" = \'страховая премия\'',
            '  AND q."Тип раздела" = \'размер выплаты\'',
            '  AND p."Раздел" = \'Сведения о страховых премиях (взносах) по договорам страхования, в разрезе страховщиков\'',
            '  AND q."Раздел" = \'Сведения о выплатах в разрезе страховщиков\'',
            '  AND p."Показатель 5" = \'Всего*\'',
            '  AND p."Показатель 1" IS NULL',
            '  AND p."Показатель 2" IS NULL',
            '  AND p."Показатель 3" IS NULL',
            '  AND p."Показатель 4" IS NULL',
            '  AND p."ед. Измерения" = q."ед. Измерения";',
            "```",
            "",
            "3. Regional comparison for one indicator:",
            "```sql",
            "SELECT",
            '  "country",',
            '  "district",',
            '  "region",',
            '  "Тип раздела",',
            '  "Раздел",',
            '  "Показатель 1",',
            '  "Показатель 2",',
            '  "Показатель 3",',
            '  "Показатель 4",',
            '  "Показатель 5",',
            '  "Значение",',
            '  "ед. Измерения"',
            'FROM "регионы"',
            'WHERE "Тип раздела" = \'страховая премия\'',
            '  AND "country" = \'На территории Российской Федерации\'',
            '  AND "district" = \'Дальневосточный федеральный округ\'',
            '  AND "Раздел" = \'Сведения о страховых премиях (взносах) по договорам страхования в разрезе федеральных округов и субъектов Российской Федерации\'',
            '  AND "Показатель 5" = \'Всего\'',
            '  AND "Показатель 1" IS NULL',
            '  AND "Показатель 2" IS NULL',
            '  AND "Показатель 3" IS NULL',
            '  AND "Показатель 4" IS NULL',
            'ORDER BY "Значение" DESC;',
            "```",
            "",
            "4. Summary-table lookup:",
            "```sql",
            "SELECT",
            '  "Тип раздела",',
            '  "Раздел",',
            '  "Показатель 1",',
            '  "Показатель 2",',
            '  "Показатель 3",',
            '  "Показатель 4",',
            '  "Показатель 5",',
            '  "Значение",',
            '  "ед. Измерения"',
            'FROM "сводные_показатели"',
            'WHERE "Раздел" = \'Страховые премии (взносы) по договорам страхования, тыс. руб.\'',
            'ORDER BY "Значение" DESC;',
            "```",
            "",
            "5. Safe aggregation pattern:",
            "```sql",
            "SELECT",
            '  "Тип раздела",',
            '  "Раздел",',
            '  "Показатель 1",',
            '  "Показатель 2",',
            '  "Показатель 3",',
            '  "Показатель 4",',
            '  "Показатель 5",',
            '  "ед. Измерения",',
            '  SUM("Значение") AS total_value',
            'FROM "<table_name>"',
            'GROUP BY "Тип раздела", "Раздел", "Показатель 1", "Показатель 2", "Показатель 3", "Показатель 4", "Показатель 5", "ед. Измерения";',
            "```",
        ]
    )

    return "\n".join(lines)


def load_table_specs(metadata_dir: Path, sqlite_path: Path) -> list[dict[str, str]]:
    if not metadata_dir.is_dir():
        raise ValueError(f"Metadata directory does not exist: {metadata_dir}")
    if not sqlite_path.is_file():
        raise ValueError(f"SQLite file does not exist: {sqlite_path}")

    json_paths = sorted(metadata_dir.glob("[0-9][0-9].*.json"))
    if not json_paths:
        raise ValueError(f"No metadata JSON files found in {metadata_dir}")

    metadata_by_subject: dict[str, list[dict]] = {
        spec["key"]: []
        for spec in LONG_SUBJECT_TABLES
    }
    for json_path in json_paths:
        metadata = json.loads(json_path.read_text(encoding="utf-8"))
        subject_key = classify_long_subject_table(metadata)
        metadata_by_subject[subject_key].append(metadata)

    with sqlite3.connect(sqlite_path) as connection:
        known_tables = {
            table_name
            for (table_name,) in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table'"
            )
        }

        table_specs: list[dict[str, str]] = []
        for spec in LONG_SUBJECT_TABLES:
            subject_key = spec["key"]
            subject_metadata = metadata_by_subject[subject_key]
            if not subject_metadata:
                continue

            table_name = build_table_name(Path(spec["file_name"]).stem)
            if table_name not in known_tables:
                raise ValueError(
                    f'Subject table "{table_name}" does not exist in SQLite database "{sqlite_path}"'
                )

            actual_columns = [
                row[1]
                for row in connection.execute(f'PRAGMA table_info("{table_name}")')
            ]
            if not actual_columns:
                raise ValueError(f'Could not read columns for SQLite table "{table_name}"')

            table_specs.append(
                {
                    "table_name": table_name,
                    "description": TABLE_DESCRIPTIONS[subject_key],
                    "row_grain": ROW_GRAINS[subject_key],
                    "identifier_columns": format_identifier_columns(actual_columns),
                    "indicator_description": INDICATOR_DESCRIPTIONS[subject_key],
                    "actual_columns": ", ".join(f'"{column}"' for column in actual_columns),
                }
            )

    return table_specs


def load_section_type_level_counts(sqlite_path: Path) -> dict[str, set[int]]:
    level_counts_by_type: dict[str, set[int]] = {}

    with sqlite3.connect(sqlite_path) as connection:
        table_names = [
            table_name
            for (table_name,) in connection.execute(
                "SELECT name FROM sqlite_master WHERE type='table' ORDER BY name"
            )
        ]

        for table_name in table_names:
            rows = connection.execute(
                f'''
                SELECT
                  "Тип раздела",
                  (
                    CASE WHEN "Показатель 1" IS NOT NULL AND "Показатель 1" != '' THEN 1 ELSE 0 END +
                    CASE WHEN "Показатель 2" IS NOT NULL AND "Показатель 2" != '' THEN 1 ELSE 0 END +
                    CASE WHEN "Показатель 3" IS NOT NULL AND "Показатель 3" != '' THEN 1 ELSE 0 END +
                    CASE WHEN "Показатель 4" IS NOT NULL AND "Показатель 4" != '' THEN 1 ELSE 0 END +
                    CASE WHEN "Показатель 5" IS NOT NULL AND "Показатель 5" != '' THEN 1 ELSE 0 END
                  ) AS indicator_depth
                FROM "{table_name}"
                WHERE "Тип раздела" IS NOT NULL
                '''
            ).fetchall()

            for section_type, indicator_depth in rows:
                if indicator_depth is None or indicator_depth <= 0:
                    continue
                level_counts_by_type.setdefault(section_type, set()).add(indicator_depth)

    return level_counts_by_type


def format_level_counts(level_counts: set[int]) -> str:
    ordered_counts = sorted(level_counts)
    return ", ".join(str(level_count) for level_count in ordered_counts)


def format_identifier_columns(actual_columns: list[str]) -> str:
    identifier_columns = [
        column
        for column in actual_columns
        if column
        not in {
            "Раздел",
            "Тип раздела",
            "Показатель 1",
            "Показатель 2",
            "Показатель 3",
            "Показатель 4",
            "Показатель 5",
            "Значение",
            "ед. Измерения",
        }
    ]
    if not identifier_columns:
        return "none"
    return ", ".join(f'"{column}"' for column in identifier_columns)


def build_table_name(stem: str) -> str:
    stem_without_numeric_prefix = re.sub(r"^\d+\.", "", stem)
    normalized_stem = "".join(
        character.lower() if character.isalnum() else "_"
        for character in stem_without_numeric_prefix
    )
    normalized_stem = re.sub(r"_+", "_", normalized_stem).strip("_")
    if not normalized_stem:
        raise ValueError(f"Cannot build table name from {stem}")
    return normalized_stem


if __name__ == "__main__":
    main()
