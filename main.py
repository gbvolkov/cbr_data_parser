from pathlib import Path
import sys

from cbr_data_parser import (
    convert_insurer_active_contract_amount_sheet_to_csv_and_json,
    convert_insurer_active_contract_count_sheet_to_csv_and_json,
    convert_insurer_contract_amount_sheet_to_csv_and_json,
    convert_insurer_contract_count_sheet_to_csv_and_json,
    convert_insurer_intermediary_electronic_premiums_sheet_to_csv_and_json,
    convert_insurer_intermediary_electronic_reward_sheet_to_csv_and_json,
    convert_insurer_intermediary_premiums_sheet_to_csv_and_json,
    convert_insurer_intermediary_reward_sheet_to_csv_and_json,
    convert_insurer_oms_sheet_to_csv_and_json,
    convert_insurer_osago_direct_reimbursement_sheet_to_csv_and_json,
    convert_insurer_osago_sheet_to_csv_and_json,
    convert_insurer_payouts_sheet_to_csv_and_json,
    convert_insurer_premiums_sheet_to_csv_and_json,
    convert_insurer_reported_claim_count_sheet_to_csv_and_json,
    convert_insurer_settled_claim_count_sheet_to_csv_and_json,
    convert_key_insurance_sheet_to_csv_and_json,
    convert_main_activity_sheet_to_csv_and_json,
    convert_ovs_members_sheet_to_csv_and_json,
    convert_regional_contract_amount_sheet_to_csv_and_json,
    convert_regional_contract_count_sheet_to_csv_and_json,
    convert_regional_payments_sheet_to_csv_and_json,
    convert_regional_premiums_sheet_to_csv_and_json,
    convert_regional_settled_claims_sheet_to_csv_and_json,
    convert_reinsurance_incoming_payments_sheet_to_csv_and_json,
    convert_reinsurance_incoming_premiums_sheet_to_csv_and_json,
    convert_reinsurance_outgoing_payments_sheet_to_csv_and_json,
    convert_reinsurance_outgoing_premiums_sheet_to_csv_and_json,
)


CONVERSIONS = [
    {
        "prefix": 1,
        "converter": convert_key_insurance_sheet_to_csv_and_json,
        "primary_label": "Показатели",
        "primary_key": "metric_rows",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 2,
        "converter": convert_main_activity_sheet_to_csv_and_json,
        "primary_label": "Показатели",
        "primary_key": "metric_rows",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 3,
        "converter": convert_insurer_premiums_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 4,
        "converter": convert_insurer_contract_count_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 5,
        "converter": convert_insurer_contract_amount_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 6,
        "converter": convert_insurer_active_contract_count_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 7,
        "converter": convert_insurer_active_contract_amount_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 8,
        "converter": convert_insurer_reported_claim_count_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 9,
        "converter": convert_insurer_settled_claim_count_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 10,
        "converter": convert_insurer_payouts_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 11,
        "converter": convert_regional_premiums_sheet_to_csv_and_json,
        "primary_label": "Регионы",
        "primary_key": "region_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 13,
        "converter": convert_regional_contract_count_sheet_to_csv_and_json,
        "primary_label": "Регионы",
        "primary_key": "region_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 15,
        "converter": convert_regional_contract_amount_sheet_to_csv_and_json,
        "primary_label": "Регионы",
        "primary_key": "region_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 16,
        "converter": convert_regional_payments_sheet_to_csv_and_json,
        "primary_label": "Регионы",
        "primary_key": "region_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 18,
        "converter": convert_regional_settled_claims_sheet_to_csv_and_json,
        "primary_label": "Регионы",
        "primary_key": "region_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 19,
        "converter": convert_reinsurance_incoming_premiums_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 20,
        "converter": convert_reinsurance_incoming_payments_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 21,
        "converter": convert_reinsurance_outgoing_premiums_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 22,
        "converter": convert_reinsurance_outgoing_payments_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 23,
        "converter": convert_insurer_osago_sheet_to_csv_and_json,
        "primary_label": "Записи",
        "primary_key": "record_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 24,
        "converter": convert_insurer_osago_direct_reimbursement_sheet_to_csv_and_json,
        "primary_label": "Записи",
        "primary_key": "record_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 25,
        "converter": convert_insurer_oms_sheet_to_csv_and_json,
        "primary_label": "Записи",
        "primary_key": "record_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
    {
        "prefix": 26,
        "converter": convert_insurer_intermediary_premiums_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 27,
        "converter": convert_insurer_intermediary_reward_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 28,
        "converter": convert_insurer_intermediary_electronic_premiums_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 29,
        "converter": convert_insurer_intermediary_electronic_reward_sheet_to_csv_and_json,
        "primary_label": "Страховщики",
        "primary_key": "insurer_row_count",
        "secondary_label": "Виды страхования",
        "secondary_key": "insurance_columns",
    },
    {
        "prefix": 30,
        "converter": convert_ovs_members_sheet_to_csv_and_json,
        "primary_label": "Записи",
        "primary_key": "record_row_count",
        "secondary_label": "Показатели",
        "secondary_key": "metric_columns",
    },
]


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")

    data_directory = Path("./data")
    output_directory = Path("./output")

    for conversion in CONVERSIONS:
        source_path = next(data_directory.glob(f"{conversion['prefix']:02d}.*.xlsx"))
        csv_path = output_directory / f"{source_path.stem}.csv"
        json_path = output_directory / f"{source_path.stem}.json"

        metadata = conversion["converter"](source_path, csv_path, json_path)

        print(f"CSV: {csv_path}")
        print(f"JSON: {json_path}")
        print(f"{conversion['primary_label']}: {_count_value(metadata[conversion['primary_key']])}")
        print(f"{conversion['secondary_label']}: {_count_value(metadata[conversion['secondary_key']])}")


def _count_value(value) -> int:
    if isinstance(value, int):
        return value
    return len(value)


if __name__ == "__main__":
    main()
