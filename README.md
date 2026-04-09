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

## Build SQLite

```powershell
uv run python csv_to_sqlite.py
```

Optional arguments:

- `--csv-dir` to choose the source directory with generated CSV files
- `--sqlite-path` to choose the output SQLite file path

The script imports every `output/[0-9][0-9].*.csv` file into one SQLite database. Each CSV becomes a separate table.

## Build Long CSV

```powershell
uv run python csv_to_long_csv.py
```

Optional arguments:

- `--csv-dir` to choose the source directory with generated wide CSV files
- `--output-dir` to choose the output directory for long CSV files

The script converts every wide CSV in `output/` into four aggregated long CSV tables in `output_long/`:

- `output_long/01.Сводные показатели.csv`
- `output_long/02.Страховщики.csv`
- `output_long/03.Регионы.csv`
- `output_long/04.Общества взаимного страхования.csv`

Each long table always contains:

- `Раздел`
- `Тип раздела`
- `Показатель 1`
- `Показатель 2`
- `Показатель 3`
- `Показатель 4`
- `Показатель 5`
- `Значение`
- `ед. Измерения`

And it additionally contains only the identifier columns that belong to the subject bucket:

- summary table: no extra identifier columns
- insurer table: insurer registration number and insurer name
- regional table: regional label plus `country`, `district`, and `region`
- OVS table: OVS registration number and insurer name

For the summary table, `Раздел` remains the original metric row.
For the subject tables, `Раздел` stores the source sheet description from metadata without the trailing unit suffix.
`Тип раздела` stores a normalized business category such as `страховая премия`, `размер выплаты`, `количество договоров`, `страховая сумма`, `ОСАГО`, or `ОМС`.
For the regional table, the original regional label is preserved and additionally split into `country`, `district`, and `region`.
Иерархия показателя больше не хранится в одной строке через `>`. Вместо этого каждый уровень записывается в отдельную колонку `Показатель 1` ... `Показатель 5`. Уровни выравниваются вправо: самый глубокий уровень всегда попадает в последнюю непустую колонку, а неиспользуемые начальные уровни остаются пустыми.
For sheets where units differ by metric column, the script derives `ед. Измерения` from the source column label; otherwise it derives the unit from metadata.

To import the long CSVs into SQLite, run:

```powershell
uv run python csv_to_sqlite.py --csv-dir output_long --sqlite-path output/cbr_data.sqlite
```

Subject grouping:

- `01`, `02` -> `Сводные показатели`
- `03`-`10`, `19`-`29` -> `Страховщики`
- `11`, `13`, `15`, `16`, `18` -> `Регионы`
- `30` -> `Общества взаимного страхования`

Implemented workbook families:

- `01`, `02`: summary sheets
- `03`-`10`: wide insurer sheets for insurance metrics
- `11`, `13`, `15`, `16`, `18`: wide regional sheets by federal districts and subjects of the Russian Federation
- `19`-`22`: wide insurer sheets for inward and outward reinsurance
- `23`, `24`, `25`, `30`: wide entity sheets for ОСАГО, ОМС, and ОВС
- `26`-`29`: wide insurer sheets for intermediary participation

Outputs are written to `output/` with the same base filename as the source workbook and both `.csv` and `.json` extensions.

CSV shapes:

- `01` and `02`: transposed; rows are metrics and columns are flattened hierarchical `Виды страхования`
- `03`-`10`, `19`-`29`: wide by insurer; first two columns are insurer identifiers, remaining columns are flattened hierarchical metric paths
- `11`, `13`, `15`, `16`, `18`: wide by region; first column is the regional label, remaining columns are flattened metric paths
- `23`, `24`, `25`, `30`: wide by entity; first two columns are entity identifiers, remaining columns are flattened metric paths

Each `.json` sidecar stores report metadata, resolved column descriptions, summary rows when present, footnotes, and skipped empty trailing columns.
