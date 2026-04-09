from pathlib import Path
import sys

from cbr_data_parser import (
    convert_insurer_active_contract_count_sheet_to_csv_and_json,
    convert_insurer_contract_amount_sheet_to_csv_and_json,
    convert_insurer_contract_count_sheet_to_csv_and_json,
    convert_insurer_premiums_sheet_to_csv_and_json,
    convert_key_insurance_sheet_to_csv_and_json,
    convert_main_activity_sheet_to_csv_and_json,
)


def main() -> None:
    sys.stdout.reconfigure(encoding="utf-8")

    source_path = Path("./data/01.Основные показатели деятельности (по ключевым видам страхования).xlsx")
    csv_path = Path("./output/01.Основные показатели деятельности (по ключевым видам страхования).csv")
    json_path = Path("./output/01.Основные показатели деятельности (по ключевым видам страхования).json")

    metadata = convert_key_insurance_sheet_to_csv_and_json(source_path, csv_path, json_path)
    print(f"CSV: {csv_path}")
    print(f"JSON: {json_path}")
    print(f"Показатели: {len(metadata['metric_rows'])}")
    print(f"Виды страхования: {len(metadata['insurance_columns'])}")

    source_path = Path("./data/02.Основные показатели деятельности.xlsx")
    csv_path = Path("./output/02.Основные показатели деятельности.csv")
    json_path = Path("./output/02.Основные показатели деятельности.json")

    metadata = convert_main_activity_sheet_to_csv_and_json(source_path, csv_path, json_path)
    print(f"CSV: {csv_path}")
    print(f"JSON: {json_path}")
    print(f"Показатели: {len(metadata['metric_rows'])}")
    print(f"Виды страхования: {len(metadata['insurance_columns'])}")

    source_path = Path("./data/03.Страховые премии по договорам страхования (в разрезе страховщиков).xlsx")
    csv_path = Path("./output/03.Страховые премии по договорам страхования (в разрезе страховщиков).csv")
    json_path = Path("./output/03.Страховые премии по договорам страхования (в разрезе страховщиков).json")

    metadata = convert_insurer_premiums_sheet_to_csv_and_json(source_path, csv_path, json_path)
    print(f"CSV: {csv_path}")
    print(f"JSON: {json_path}")
    print(f"Страховщики: {metadata['insurer_row_count']}")
    print(f"Виды страхования: {len(metadata['insurance_columns'])}")

    source_path = Path("./data/04.Количество заключенных договоров (в разрезе страховщиков).xlsx")
    csv_path = Path("./output/04.Количество заключенных договоров (в разрезе страховщиков).csv")
    json_path = Path("./output/04.Количество заключенных договоров (в разрезе страховщиков).json")

    metadata = convert_insurer_contract_count_sheet_to_csv_and_json(source_path, csv_path, json_path)
    print(f"CSV: {csv_path}")
    print(f"JSON: {json_path}")
    print(f"Страховщики: {metadata['insurer_row_count']}")
    print(f"Виды страхования: {len(metadata['insurance_columns'])}")

    source_path = Path("./data/05.Страховые суммы по заключенным договорам (в разрезе страховщиков).xlsx")
    csv_path = Path("./output/05.Страховые суммы по заключенным договорам (в разрезе страховщиков).csv")
    json_path = Path("./output/05.Страховые суммы по заключенным договорам (в разрезе страховщиков).json")

    metadata = convert_insurer_contract_amount_sheet_to_csv_and_json(source_path, csv_path, json_path)
    print(f"CSV: {csv_path}")
    print(f"JSON: {json_path}")
    print(f"Страховщики: {metadata['insurer_row_count']}")
    print(f"Виды страхования: {len(metadata['insurance_columns'])}")

    source_path = Path("./data/06.Количество действовавших договоров (в разрезе страховщиков).xlsx")
    csv_path = Path("./output/06.Количество действовавших договоров (в разрезе страховщиков).csv")
    json_path = Path("./output/06.Количество действовавших договоров (в разрезе страховщиков).json")

    metadata = convert_insurer_active_contract_count_sheet_to_csv_and_json(
        source_path, csv_path, json_path
    )
    print(f"CSV: {csv_path}")
    print(f"JSON: {json_path}")
    print(f"Страховщики: {metadata['insurer_row_count']}")
    print(f"Виды страхования: {len(metadata['insurance_columns'])}")


if __name__ == "__main__":
    main()
