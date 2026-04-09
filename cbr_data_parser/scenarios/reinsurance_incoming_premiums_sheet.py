from __future__ import annotations

from pathlib import Path

from cbr_data_parser.utils.reinsurance_wide_sheet import (
    convert_reinsurance_wide_sheet_to_csv_and_json,
)


EXPECTED_TITLE = "Сведения о страховых премиях по договорам входящего перестрахования в разрезе страховщиков, тыс руб."


def convert_reinsurance_incoming_premiums_sheet_to_csv_and_json(
    excel_path: Path | str,
    csv_path: Path | str,
    json_path: Path | str,
) -> dict:
    return convert_reinsurance_wide_sheet_to_csv_and_json(
        excel_path=excel_path,
        csv_path=csv_path,
        json_path=json_path,
        expected_title=EXPECTED_TITLE,
        report_date_row=3,
        report_period_row=4,
        id_header_row=5,
        header_row_start=5,
        header_row_end=8,
        summary_row_index=9,
        expected_data_column_count=64,
    )
