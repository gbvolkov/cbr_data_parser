from __future__ import annotations


LONG_SUBJECT_TABLES = [
    {
        "key": "summary",
        "file_name": "01.Сводные показатели.csv",
    },
    {
        "key": "insurers",
        "file_name": "02.Страховщики.csv",
    },
    {
        "key": "regions",
        "file_name": "03.Регионы.csv",
    },
    {
        "key": "mutual_societies",
        "file_name": "04.Общества взаимного страхования.csv",
    },
]

REGIONAL_DIMENSION_HEADERS = ["country", "district", "region"]
REGIONAL_COUNTRY_LABELS = {
    "На территории Российской Федерации",
    "За пределами Российской Федерации",
}

SECTION_TYPE_DESCRIPTIONS = {
    "страховая премия": "Premium-related sections, including direct insurance premiums, reinsurance premiums, and premium income references.",
    "размер выплаты": "Payout-related sections, including insurance payouts, court-ordered payouts, and payout expense references.",
    "количество договоров": "Sections about the count of insurance contracts, both concluded and active.",
    "страховая сумма": "Sections about insured amounts or total sums under insurance contracts.",
    "число застрахованных": "Sections about the number of insured persons.",
    "количество страховых случаев": "Sections about reported or processed insurance case counts in mixed summary blocks.",
    "количество заявленных страховых случаев": "Sections specifically about reported insurance case counts.",
    "количество урегулированных страховых случаев": "Sections specifically about settled insurance case counts.",
    "неурегулированный страховой случай": "Sections about outstanding unresolved insurance cases and their amounts.",
    "возврат страховой премии": "Sections about returned premiums and the related contract counts.",
    "выкупная сумма": "Sections about surrender values and related contract counts.",
    "негарантированная выплата": "Sections about non-guaranteed payouts and their counts.",
    "дополнительная выплата": "Sections about additional payouts and their counts.",
    "возмещение расходов страхователя": "Sections about reimbursed policyholder expenses aimed at loss reduction.",
    "неустойка": "Sections about penalties for delayed insurance payout processing.",
    "вознаграждение посредникам": "Sections about intermediary remuneration and commissions.",
    "ОСАГО": "Operational and financial sections from the ОСАГО subject table.",
    "прямое возмещение убытков по ОСАГО": "Sections about direct loss compensation flows within ОСАГО.",
    "ОМС": "Operational sections from the mandatory medical insurance subject table.",
    "число членов ОВС": "Sections about mutual insurance society membership counts and changes.",
}


def classify_long_subject_table(metadata: dict) -> str:
    if metadata.get("row_id_column"):
        return "summary"

    if metadata.get("row_label_column"):
        return "regions"

    id_columns = metadata.get("id_columns", [])
    if not id_columns:
        raise ValueError("Metadata does not contain row identifiers.")

    first_label = str(id_columns[0]["label"]).lower()
    if "общества взаимного страхования" in first_label:
        return "mutual_societies"
    if "страховщика" in first_label:
        return "insurers"

    raise ValueError(f"Unsupported subject identifiers: {id_columns}")


def resolve_section_type(section: str) -> str:
    normalized = section.lower()

    if "прямом возмещении убытков" in normalized:
        return "прямое возмещение убытков по ОСАГО"
    if "обязательном медицинском страховании" in normalized:
        return "ОМС"
    if "обязательном страховании гражданской ответственности владельцев транспортных средств" in normalized:
        return "ОСАГО"
    if "членов общества взаимного страхования" in normalized:
        return "число членов ОВС"
    if "вознаграждении посредникам" in normalized:
        return "вознаграждение посредникам"
    if "возврат страховых премий" in normalized or "сумма возврата страховых премий" in normalized:
        return "возврат страховой премии"
    if "неустойка" in normalized:
        return "неустойка"
    if "возмещение расходов страхователей" in normalized:
        return "возмещение расходов страхователя"
    if "дополнительные выплаты" in normalized:
        return "дополнительная выплата"
    if "негарантированные выплаты" in normalized:
        return "негарантированная выплата"
    if "выкупные суммы" in normalized:
        return "выкупная сумма"
    if "неурегулированные страховые случаи" in normalized:
        return "неурегулированный страховой случай"
    if "урегулированных страховых случаев" in normalized:
        return "количество урегулированных страховых случаев"
    if "заявленных страховых случаев" in normalized:
        return "количество заявленных страховых случаев"
    if "количество страховых случаев" in normalized:
        return "количество страховых случаев"
    if "число застрахованных" in normalized:
        return "число застрахованных"
    if "количестве договоров" in normalized or "количество договоров" in normalized:
        return "количество договоров"
    if "страховых суммах" in normalized or "страховая сумма" in normalized:
        return "страховая сумма"
    if (
        "страховых выплатах" in normalized
        or "страховые выплаты" in normalized
        or "cтраховых выплатах" in normalized
        or "cтраховые выплаты" in normalized
        or "сведения о выплатах" in normalized
        or "выплатах по договорам страхования" in normalized
    ):
        return "размер выплаты"
    if normalized.startswith("выплаты >") or normalized.startswith("выплаты по договорам страхования"):
        return "размер выплаты"
    if "сумма выплат" in normalized or "сумма страховых выплат" in normalized:
        return "размер выплаты"
    if "страховых премиях" in normalized or "страховые премии" in normalized:
        return "страховая премия"
    if "премии (доходы)" in normalized:
        return "страховая премия"

    raise ValueError(f'Unsupported section type for "{section}"')


def resolve_regional_dimensions(
    row_label: str,
    hierarchy_state: dict[str, str | None],
) -> list[str]:
    if row_label in REGIONAL_COUNTRY_LABELS:
        hierarchy_state["country"] = row_label
        hierarchy_state["district"] = None
        return [row_label, "", ""]

    country = hierarchy_state["country"]
    if country is None:
        raise ValueError(f'Regional row "{row_label}" appears before country label.')

    if row_label.endswith("федеральный округ"):
        hierarchy_state["district"] = row_label
        return [country, row_label, ""]

    district = hierarchy_state["district"]
    if country == "На территории Российской Федерации" and district is None:
        raise ValueError(f'Regional row "{row_label}" appears before district label.')

    return [country, district or "", row_label]
