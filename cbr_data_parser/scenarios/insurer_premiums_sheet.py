from __future__ import annotations

from pathlib import Path

from cbr_data_parser.utils import convert_insurer_wide_sheet_to_csv_and_json


EXPECTED_TITLE = "Сведения о страховых премиях (взносах) по договорам страхования, в разрезе страховщиков, тыс руб."
EXPECTED_DATA_COLUMN_COUNT = 256


def convert_insurer_premiums_sheet_to_csv_and_json(
    excel_path: Path | str,
    csv_path: Path | str,
    json_path: Path | str,
) -> dict:
    return convert_insurer_wide_sheet_to_csv_and_json(
        excel_path=excel_path,
        csv_path=csv_path,
        json_path=json_path,
        expected_title=EXPECTED_TITLE,
        expected_data_column_count=EXPECTED_DATA_COLUMN_COUNT,
    )
