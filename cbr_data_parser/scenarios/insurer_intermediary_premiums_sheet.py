from __future__ import annotations

from pathlib import Path

from cbr_data_parser.utils.insurer_intermediary_participation_sheet import (
    convert_insurer_intermediary_participation_sheet_to_csv_and_json,
)


EXPECTED_TITLE = "Сведения о страховых премиях (взносах) по договорам страхования, заключенным при участии посредников (кроме договоров страхования, заключенных путем обмена информацией в электронной форме), тыс руб."


def convert_insurer_intermediary_premiums_sheet_to_csv_and_json(
    excel_path: Path | str,
    csv_path: Path | str,
    json_path: Path | str,
) -> dict:
    return convert_insurer_intermediary_participation_sheet_to_csv_and_json(
        excel_path=excel_path,
        csv_path=csv_path,
        json_path=json_path,
        expected_title=EXPECTED_TITLE,
    )
