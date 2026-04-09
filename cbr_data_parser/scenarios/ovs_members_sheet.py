from __future__ import annotations

from pathlib import Path

from cbr_data_parser.utils.wide_entity_sheet import convert_wide_entity_sheet_to_csv_and_json


EXPECTED_TITLE = "Сведения о числе членов общества взаимного страхования, ед."
EXPECTED_ID_LABEL_A = "Регистрационный номер записи общества взаимного страхования в едином государственном реестре субъектов страхового дела"
EXPECTED_ID_LABEL_B = "Полное наименование страховщика"


def convert_ovs_members_sheet_to_csv_and_json(
    excel_path: Path | str,
    csv_path: Path | str,
    json_path: Path | str,
) -> dict:
    return convert_wide_entity_sheet_to_csv_and_json(
        excel_path,
        csv_path,
        json_path,
        expected_title=EXPECTED_TITLE,
        expected_id_label_a=EXPECTED_ID_LABEL_A,
        expected_id_label_b=EXPECTED_ID_LABEL_B,
        header_row_start=4,
        header_row_end=6,
        summary_row_index=7,
        summary_label="Итого",
        data_start_row=8,
        expected_metric_count=9,
        metric_start_column=3,
    )
