# cbr_data_parser

## Development rules

1. Keep every solution simple and explicit. Do not introduce unnecessary abstractions, classes, or indirection.
2. Do not add fallbacks or workarounds. If the environment is broken or the input data is invalid, report the error and stop.
3. Preserve layered architecture. Keep scenario code separate from utility code.
4. Prefer UTF-8 whenever possible. Do not use Unicode escape sequences such as `\uXXXX` unless there is a hard technical requirement.

## Requirements

- Python 3.13
- `uv`

## Environment setup

```powershell
uv sync
```

## Run

```powershell
uv run python main.py
```

This converts:

- `data/01.Основные показатели деятельности (по ключевым видам страхования).xlsx`
- `data/02.Основные показатели деятельности.xlsx`
- `data/03.Страховые премии по договорам страхования (в разрезе страховщиков).xlsx`
- `data/04.Количество заключенных договоров (в разрезе страховщиков).xlsx`
- `data/05.Страховые суммы по заключенным договорам (в разрезе страховщиков).xlsx`
- `data/06.Количество действовавших договоров (в разрезе страховщиков).xlsx`

into:

- `output/01.Основные показатели деятельности (по ключевым видам страхования).csv`
- `output/01.Основные показатели деятельности (по ключевым видам страхования).json`
- `output/02.Основные показатели деятельности.csv`
- `output/02.Основные показатели деятельности.json`
- `output/03.Страховые премии по договорам страхования (в разрезе страховщиков).csv`
- `output/03.Страховые премии по договорам страхования (в разрезе страховщиков).json`
- `output/04.Количество заключенных договоров (в разрезе страховщиков).csv`
- `output/04.Количество заключенных договоров (в разрезе страховщиков).json`
- `output/05.Страховые суммы по заключенным договорам (в разрезе страховщиков).csv`
- `output/05.Страховые суммы по заключенным договорам (в разрезе страховщиков).json`
- `output/06.Количество действовавших договоров (в разрезе страховщиков).csv`
- `output/06.Количество действовавших договоров (в разрезе страховщиков).json`

For `01` and `02`, the CSV is transposed:

- rows are metrics (`Показатель`)
- columns are flattened hierarchical `Виды страхования`

For `03`, `04`, `05`, and `06`, the CSV stays wide by insurer:

- rows are insurers
- the first two columns are insurer identifiers
- the remaining columns are flattened hierarchical `Виды страхования`

The JSON sidecar stores report metadata and the resolved hierarchy paths.
