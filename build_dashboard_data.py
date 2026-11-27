import argparse
import json
import math
import warnings
from pathlib import Path

from openpyxl import load_workbook


PRODUCTS = [
    "Mobile",
    "Display",
    "Climate",
    "Others",
    "Non Standard",
    "Blank",
]
REQUIRED_PRODUCTS = PRODUCTS[:-1]
REGIONS = ["SEUK SMB", "SEBN", "SEG", "SESA", "SEF", "SEP", "SEH", "SEACE"]
MONTHS = [
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
]
COUNTRY_ALIASES = {"beligum": "Belgium"}

METRIC_SHEETS = {
    "Total Database": ("db", "Total Database"),
    "Total - Opt-in contacts": ("optIn", "Total - Opt-in Contacts"),
    "Total - Opt-Out contacts": ("optOut", "Total - Opt-Out Contacts"),
    "Total - Opt-in-out Contacts": (
        "neither",
        "Total - Niether Opt-in/Out Contacts",
    ),
    "New Contacts": ("newContacts", "New Contacts"),
    "New Subscribers": ("newSubscribers", "New Subscribers"),
    "New UnSubscribers": ("newUnsubscribers", "New UnSubscribers"),
    "Marketable Database": ("market", "Marketable Database"),
}
ESTORE_SHEETS = {
    "Estore Contacts - Registered": (
        "registered",
        "Estore Contacts - Registered",
    ),
    "Estore Contacts - SALE ": ("sale", "Estore Contacts - Sale"),
    "Estore Contacts - Abandoned": (
        "abandoned",
        "Estore Contacts - Abandoned",
    ),
}
ENGAGEMENT_SHEETS = {
    "Engaged Contacts": ("engaged", "Engaged Contacts"),
    "Touched Contacts": ("touched", "Touched Contacts last 3 month"),
}
MARKETABLE_BREAKDOWN_SHEETS = {
    "Marketable - I4B": (
        "i4b",
        "Marketable Database - I4B",
    ),
    "Marketable - Without I4B": (
        "withoutI4B",
        " Marketable Database - Without I4B",
    ),
    "Marketable - Non Estore Contact": (
        "nonEstore",
        "Marketable - Non Estore Contacts",
    ),
}
AUDIENCE_SHEETS = {
    "Marketable - Global Industry": (
        "industry",
        "Global Industry",
        "Marketable - Global Industry Count",
    ),
    "Marketable -No of employee": (
        "employees",
        "No of employee",
        "Marketable - No of employee Count",
    ),
    "Marketable - Lead Source": (
        "sources",
        "Lead_Source",
        "Marketable",
    ),
}
CAMPAIGN_SHEET = "Campaign Performance"
FORM_SHEET = "Form submission"


def normalized(value):
    return " ".join(str(value or "").strip().lower().split())


def country_name(value):
    name = str(value or "").strip()
    return COUNTRY_ALIASES.get(name.casefold(), name)


def number(value, location, blank_as_zero=False):
    if value is None or value == "":
        if blank_as_zero:
            return 0
        raise ValueError(f"{location}: expected a number, found a blank cell.")
    if isinstance(value, str):
        value = value.replace(",", "").strip()
    try:
        result = float(value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{location}: expected a number, found {value!r}.") from error
    if not math.isfinite(result) or result < 0 or not result.is_integer():
        raise ValueError(f"{location}: invalid count {value!r}.")
    return int(result)


def optional_number(value, location):
    if value is None or value == "":
        return None
    if isinstance(value, str) and value.strip().upper() in {"NA", "N/A", "BLANK"}:
        return None
    return number(value, location)


def rows_for_metric(workbook, sheet_name, metric_name):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the required sheet {sheet_name!r}.")
    worksheet = workbook[sheet_name]
    iterator = worksheet.iter_rows(values_only=True)
    headers = [normalized(value) for value in next(iterator)]
    header_indexes = {name: index for index, name in enumerate(headers) if name}
    metric_index = header_indexes.get(normalized(metric_name))
    if metric_index is None:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing the {metric_name!r} column."
        )
    product_indexes = []
    for product in REQUIRED_PRODUCTS:
        if normalized(product) not in header_indexes:
            raise ValueError(
                f"Sheet {sheet_name!r} is missing the {product!r} column."
            )
    for product in PRODUCTS:
        product_indexes.append(
            headers.index(normalized(product))
            if normalized(product) in headers
            else None
        )
    quarter_header = "quarter" if "quarter" in header_indexes else "quater"
    key_indexes = [
        header_indexes.get(name)
        for name in ("year", quarter_header, "month", "region")
    ]
    if any(index is None for index in key_indexes):
        raise ValueError(
            f"Sheet {sheet_name!r} is missing a year, quarter, month, or region column."
        )
    country_index = header_indexes.get("country")

    result = {}
    country_result = {}
    summary_keys = set()
    for row_number, row in enumerate(iterator, start=2):
        key_cells = tuple(
            row[index] if index < len(row) else None for index in key_indexes
        )
        value = optional_number(
            row[metric_index] if metric_index < len(row) else None,
            f"{sheet_name!r} row {row_number}",
        )
        product_values = [
            optional_number(
                row[index] if index < len(row) else None,
                f"{sheet_name!r} row {row_number}",
            )
            if index is not None
            else None
            for index in product_indexes
        ]
        if value is None and all(cell is None for cell in product_values):
            continue
        if any(cell is None or cell == "" for cell in key_cells):
            raise ValueError(f"Sheet {sheet_name!r}, row {row_number}: missing key.")
        year, quarter, month, region = [str(cell).strip() for cell in key_cells]
        country = country_name(
            row[country_index]
            if country_index is not None and country_index < len(row)
            else ""
        )
        if region not in REGIONS and country in REGIONS:
            country, region = region, country
        if month not in MONTHS or region not in REGIONS:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown month or region."
            )
        if quarter not in {"Q1", "Q2", "Q3", "Q4"}:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown quarter {quarter!r}."
            )
        key = (year, quarter, month, region)
        if country and country.lower() != "all":
            country_key = (*key, country)
            country_record = country_result.setdefault(
                country_key,
                {"value": None, "products": [None] * len(PRODUCTS)},
            )
            if value is not None:
                country_record["value"] = (country_record["value"] or 0) + value
            for index, product_value in enumerate(product_values):
                if product_value is not None:
                    country_record["products"][index] = (
                        (country_record["products"][index] or 0) + product_value
                    )
        if key in summary_keys and country.lower() != "all":
            continue
        if country.lower() == "all":
            result[key] = {"value": value, "products": product_values}
            summary_keys.add(key)
            continue
        record = result.setdefault(
            key,
            {"value": None, "products": [None] * len(PRODUCTS)},
        )
        if value is not None:
            record["value"] = (record["value"] or 0) + value
        for index, product_value in enumerate(product_values):
            if product_value is not None:
                record["products"][index] = (
                    (record["products"][index] or 0) + product_value
                )
    return result, country_result


def rows_for_dimension(workbook, sheet_name, category_name, metric_name):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the required sheet {sheet_name!r}.")
    worksheet = workbook[sheet_name]
    iterator = worksheet.iter_rows(values_only=True)
    headers = [normalized(value) for value in next(iterator)]
    header_indexes = {name: index for index, name in enumerate(headers) if name}
    quarter_header = "quarter" if "quarter" in header_indexes else "quater"
    metric_header = normalized(metric_name)
    if normalized(category_name) == "lead_source":
        metric_header = next(
            (
                header
                for header in (
                    metric_header,
                    normalized("Marketable-Lead_Source count"),
                    normalized("Lead_Source count"),
                )
                if header in header_indexes
            ),
            metric_header,
        )
    required = [
        "year",
        quarter_header,
        "month",
        "region",
        normalized(category_name),
        metric_header,
    ]
    missing = [name for name in required if name not in header_indexes]
    if missing:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing columns: {', '.join(missing)}."
        )
    missing_products = [
        normalized(product)
        for product in REQUIRED_PRODUCTS
        if normalized(product) not in header_indexes
    ]
    if missing_products:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing product columns: "
            f"{', '.join(missing_products)}."
        )

    indexes = [header_indexes[name] for name in required]
    product_indexes = [
        headers.index(normalized(product))
        if normalized(product) in headers
        else None
        for product in PRODUCTS
    ]
    result = []
    for row_number, row in enumerate(iterator, start=2):
        cells = [row[index] if index < len(row) else None for index in indexes]
        year, quarter, month, region, category, value = cells
        value = optional_number(value, f"{sheet_name!r} row {row_number}")
        products = [
            optional_number(row[index], f"{sheet_name!r} row {row_number}")
            if index is not None and index < len(row)
            else None
            for index in product_indexes
        ]
        if value is None and all(product is None for product in products):
            continue
        key_values = (year, quarter, month, region)
        if all(value is None or value == "" for value in key_values):
            warnings.warn(
                f"Skipped {sheet_name!r} row {row_number}: metric values have no date or region key.",
                stacklevel=2,
            )
            continue
        if any(value is None or value == "" for value in key_values):
            raise ValueError(f"Sheet {sheet_name!r}, row {row_number}: missing key.")
        country_index = header_indexes.get("country")
        country = country_name(
            row[country_index]
            if country_index is not None and country_index < len(row)
            else ""
        )
        if region not in REGIONS and country in REGIONS:
            country, region = region, country
        year, quarter, month, region = [
            str(value).strip() for value in (year, quarter, month, region)
        ]
        if month not in MONTHS or region not in REGIONS:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown month or region."
            )
        if quarter not in {"Q1", "Q2", "Q3", "Q4"}:
            raise ValueError(
                f"Sheet {sheet_name!r}, row {row_number}: unknown quarter {quarter!r}."
            )
        record = {
                "year": year,
                "quarter": quarter,
                "month": month,
                "region": region,
                "category": str(category or "Blank").strip() or "Blank",
                "value": value,
                "products": products,
            }
        summary_key = (year, quarter, month, region, record["category"])
        result.append((summary_key, country, record))
    region_totals = {}
    summary_rows = {}
    country_rows = []
    for key, country, record in result:
        if country.lower() == "all":
            summary_rows[key] = record
            continue
        region_total = region_totals.setdefault(
            key,
            {
                **record,
                "value": None,
                "products": [None] * len(PRODUCTS),
            },
        )
        if record["value"] is not None:
            region_total["value"] = (region_total["value"] or 0) + record["value"]
        for index, value in enumerate(record["products"]):
            if value is not None:
                region_total["products"][index] = (
                    (region_total["products"][index] or 0) + value
                )
        if country:
            country_rows.append({**record, "country": country})

    aggregate_rows = []
    for key in dict.fromkeys((*region_totals, *summary_rows)):
        aggregate = summary_rows.get(key, region_totals.get(key))
        aggregate_rows.append(
            {
                **aggregate,
                "products": list(aggregate["products"]),
                "country": None,
            }
        )
    return aggregate_rows + country_rows


def cell_value(row, indexes, key):
    index = indexes.get(key)
    return row[index] if index is not None and index < len(row) else None


def nullable_number(value, location):
    if value is None or value == "" or (
        isinstance(value, str) and value.strip().upper() in {"NA", "N/A", "BLANK"}
    ):
        return None
    try:
        result = float(value.replace(",", "").strip() if isinstance(value, str) else value)
    except (TypeError, ValueError) as error:
        raise ValueError(f"{location}: expected a number, found {value!r}.") from error
    if not math.isfinite(result):
        raise ValueError(f"{location}: invalid number {value!r}.")
    return result


def data_sheet_status(worksheet, metric_names):
    rows = list(worksheet.iter_rows(values_only=True))
    headers = [normalized(value) for value in rows[0]] if rows else []
    indexes = {name: index for index, name in enumerate(headers) if name}
    metric_index = next(
        (indexes[normalized(name)] for name in metric_names if normalized(name) in indexes),
        None,
    )
    product_indexes = [
        headers.index(normalized(product))
        for product in PRODUCTS
        if normalized(product) in headers
    ]
    value_indexes = (
        ([metric_index] if metric_index is not None else []) + product_indexes
    )
    values = [
        row[metric_index]
        for row in rows[1:]
        if metric_index is not None
        and metric_index < len(row)
        and isinstance(row[metric_index], (int, float))
    ]
    if not values:
        values = [
            row[index]
            for row in rows[1:]
            for index in value_indexes
            if index < len(row) and isinstance(row[index], (int, float))
        ]
    if not values:
        return "blank"
    if len(set(values)) <= max(12, math.floor(len(values) * 0.1)):
        return "included-placeholder"
    return "loaded"


def read_detail_sheet(workbook, sheet_name, kind):
    if sheet_name not in workbook.sheetnames:
        raise ValueError(f"Workbook is missing the required sheet {sheet_name!r}.")
    worksheet = workbook[sheet_name]
    iterator = worksheet.iter_rows(values_only=True)
    headers = [normalized(value) for value in next(iterator)]
    indexes = {name: index for index, name in enumerate(headers) if name}
    quarter_header = "quarter" if "quarter" in indexes else "quater"
    campaign_name_header = (
        "campaign name"
        if "campaign name" in indexes
        else "campaign / program name"
    )
    required = (
        {
            "year",
            quarter_header,
            "month",
            "region",
            "type",
            campaign_name_header,
            "total_sent",
            "total_delivered",
            "total_clicks",
            "unsubscribes",
        }
        if kind == "campaign"
        else {
            "year",
            quarter_header,
            "month",
            "region",
            "form name",
            "type",
            "total submission",
            "unique submission",
            "new contacts",
            "marketable contacts",
        }
    )
    missing = required - indexes.keys()
    if missing:
        raise ValueError(
            f"Sheet {sheet_name!r} is missing columns: {', '.join(sorted(missing))}."
        )

    result = []
    for row_number, row in enumerate(iterator, start=2):
        if not any(value is not None and value != "" for value in row):
            continue
        get = lambda key: cell_value(row, indexes, key)
        if kind == "campaign":
            result.append(
                {
                    "year": str(get("year") or "").strip(),
                    "quarter": str(get(quarter_header) or "").strip(),
                    "month": str(get("month") or "").strip(),
                    "region": str(get("region") or "").strip(),
                    "country": country_name(get("country")) or None,
                    "type": str(get("type") or "").strip(),
                    "campaign": str(get(campaign_name_header) or "").strip(),
                    "sent": nullable_number(get("total_sent"), f"{sheet_name} row {row_number}"),
                    "delivered": nullable_number(get("total_delivered"), f"{sheet_name} row {row_number}"),
                    "uniqueOpens": nullable_number(get("unique_html_opens"), f"{sheet_name} row {row_number}"),
                    "clicks": nullable_number(get("total_clicks"), f"{sheet_name} row {row_number}"),
                    "uniqueClicks": nullable_number(get("unique_clicks"), f"{sheet_name} row {row_number}"),
                    "unsubscribes": nullable_number(get("unsubscribes"), f"{sheet_name} row {row_number}"),
                }
            )
        else:
            result.append(
                {
                    "year": str(get("year") or "").strip(),
                    "quarter": str(get(quarter_header) or "").strip(),
                    "month": str(get("month") or "").strip(),
                    "region": str(get("region") or "").strip(),
                    "country": country_name(get("country")) or None,
                    "name": str(get("form name") or "").strip(),
                    "type": str(get("type") or "").strip(),
                    "submissions": nullable_number(get("total submission"), f"{sheet_name} row {row_number}"),
                    "uniqueSubmissions": nullable_number(get("unique submission"), f"{sheet_name} row {row_number}"),
                    "newContacts": nullable_number(get("new contacts"), f"{sheet_name} row {row_number}"),
                    "marketableContacts": nullable_number(get("marketable contacts"), f"{sheet_name} row {row_number}"),
                }
            )
    return result


def audit_sheets(workbook):
    audit = []
    loaded_sheets = (
        set(METRIC_SHEETS)
        | set(ESTORE_SHEETS)
        | set(ENGAGEMENT_SHEETS)
        | set(MARKETABLE_BREAKDOWN_SHEETS)
        | set(AUDIENCE_SHEETS)
        | {CAMPAIGN_SHEET, FORM_SHEET}
    )
    for worksheet in workbook.worksheets:
        rows = list(worksheet.iter_rows(values_only=True))
        headers = [normalized(value) for value in rows[0]] if rows else []
        populated_rows = [
            row for row in rows[1:] if any(value is not None and value != "" for value in row)
        ]
        if worksheet.title in loaded_sheets:
            if worksheet.title in MARKETABLE_BREAKDOWN_SHEETS:
                status = data_sheet_status(
                    worksheet,
                    [MARKETABLE_BREAKDOWN_SHEETS[worksheet.title][1]],
                )
            elif worksheet.title in ESTORE_SHEETS:
                status = data_sheet_status(
                    worksheet,
                    [ESTORE_SHEETS[worksheet.title][1]],
                )
            elif worksheet.title in ENGAGEMENT_SHEETS:
                status = data_sheet_status(
                    worksheet,
                    [ENGAGEMENT_SHEETS[worksheet.title][1]],
                )
            elif worksheet.title in AUDIENCE_SHEETS:
                status = data_sheet_status(
                    worksheet,
                    [
                        AUDIENCE_SHEETS[worksheet.title][2],
                        "Marketable-Lead_Source count",
                        "Lead_Source count",
                    ],
                )
            else:
                status = "loaded"
        else:
            metric_index = 5 if len(headers) > 5 else 4
            values = [
                row[metric_index]
                for row in populated_rows
                if metric_index < len(row) and isinstance(row[metric_index], (int, float))
            ]
            if not values:
                status = "blank"
            elif len(set(values)) <= max(12, math.floor(len(values) * 0.1)):
                status = "placeholder"
            else:
                status = "not-integrated"
        audit.append(
            {
                "name": worksheet.title,
                "populatedRows": len(populated_rows),
                "status": status,
            }
        )
    return audit


def metric_records(rows):
    return [
        {
            "year": key[0],
            "quarter": key[1],
            "month": key[2],
            "region": key[3],
            "country": key[4] if len(key) > 4 else None,
            "value": values["value"],
            "products": values["products"],
        }
        for key, values in rows.items()
    ]


def build_payload(input_path):
    workbook = load_workbook(input_path, read_only=True, data_only=True)
    metric_rows = {}
    country_metric_rows = {}
    for sheet_name, (metric, column_name) in METRIC_SHEETS.items():
        metric_rows[metric], country_metric_rows[metric] = rows_for_metric(
            workbook,
            sheet_name,
            column_name,
        )
    estore_rows = {}
    country_estore_rows = {}
    for sheet_name, (metric, column_name) in ESTORE_SHEETS.items():
        estore_rows[metric], country_estore_rows[metric] = rows_for_metric(
            workbook,
            sheet_name,
            column_name,
        )
    engagement_rows = {}
    country_engagement_rows = {}
    for sheet_name, (metric, column_name) in ENGAGEMENT_SHEETS.items():
        engagement_rows[metric], country_engagement_rows[metric] = rows_for_metric(
            workbook,
            sheet_name,
            column_name,
        )
    marketable_breakdown = {}
    country_marketable_breakdown = {}
    for sheet_name, (metric, column_name) in MARKETABLE_BREAKDOWN_SHEETS.items():
        (
            marketable_breakdown[metric],
            country_marketable_breakdown[metric],
        ) = rows_for_metric(
            workbook,
            sheet_name,
            column_name,
        )
    audience_rows = {}
    for sheet_name, (dimension, category_column, metric_column) in AUDIENCE_SHEETS.items():
        audience_rows[dimension] = rows_for_dimension(
            workbook,
            sheet_name,
            category_column,
            metric_column,
        )

    records = {}
    for metric, rows in metric_rows.items():
        for key, values in rows.items():
            record = records.setdefault(
                key,
                {
                    "year": key[0],
                    "q": key[1],
                    "m": key[2],
                    "r": key[3],
                    "country": None,
                },
            )
            records[key][metric] = values["value"]
            records[key][f"{metric}Products"] = values["products"]

    country_records = {}
    for metric, rows in country_metric_rows.items():
        for key, values in rows.items():
            record = country_records.setdefault(
                key,
                {
                    "year": key[0],
                    "q": key[1],
                    "m": key[2],
                    "r": key[3],
                    "country": key[4],
                },
            )
            record[metric] = values["value"]
            record[f"{metric}Products"] = values["products"]

    for record in records.values():
        db_products = record.get("dbProducts", [])
        if (
            record.get("db") is not None
            and db_products
            and all(value is not None for value in db_products)
            and sum(db_products) != record["db"]
        ):
            raise ValueError(
                f"Product totals do not match Total Database for "
                f"{record['year']}, {record['m']}, {record['r']}."
            )

    metrics = sorted(metric_rows.keys())
    estore_records = {}
    for metric, rows in estore_rows.items():
        for key, values in rows.items():
            record = estore_records.setdefault(
                key,
                {
                    "year": key[0],
                    "quarter": key[1],
                    "month": key[2],
                    "region": key[3],
                    "country": None,
                },
            )
            record[metric] = values["value"]
            record[f"{metric}Products"] = values["products"]

    country_estore_records = {}
    for metric, rows in country_estore_rows.items():
        for key, values in rows.items():
            record = country_estore_records.setdefault(
                key,
                {
                    "year": key[0],
                    "quarter": key[1],
                    "month": key[2],
                    "region": key[3],
                    "country": key[4],
                },
            )
            record[metric] = values["value"]
            record[f"{metric}Products"] = values["products"]
    years = {record["year"] for record in records.values()}
    expected_records = len(years) * len(MONTHS) * len(REGIONS)
    return {
        "sourceWorkbook": input_path.name,
        "metrics": metrics,
        "records": list(records.values()) + list(country_records.values()),
        "estoreMetrics": sorted(estore_rows.keys()),
        "estore": list(estore_records.values())
        + list(country_estore_records.values()),
        "engagementMetrics": sorted(engagement_rows.keys()),
        "engagement": {
            metric: metric_records(rows)
            + metric_records(country_engagement_rows[metric])
            for metric, rows in engagement_rows.items()
        },
        "marketableBreakdownMetrics": sorted(marketable_breakdown.keys()),
        "marketableBreakdown": {
            metric: metric_records(rows)
            + metric_records(country_marketable_breakdown[metric])
            for metric, rows in marketable_breakdown.items()
        },
        "audienceDimensions": audience_rows,
        "missingDatabaseRows": max(
            0,
            expected_records
            - sum(
                values["value"] is not None
                for values in metric_rows["db"].values()
            ),
        ),
        "campaigns": read_detail_sheet(workbook, CAMPAIGN_SHEET, "campaign"),
        "forms": read_detail_sheet(workbook, FORM_SHEET, "form"),
        "sheetAudit": audit_sheets(workbook),
    }


def main():
    parser = argparse.ArgumentParser(
        description="Convert the Samsung QBR workbook into data the static dashboard can load."
    )
    parser.add_argument("workbook", type=Path, help="Path to the updated .xlsx workbook")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path(__file__).with_name("dashboard-data.json"),
        help="Output JSON path (defaults to dashboard-data.json beside this script)",
    )
    args = parser.parse_args()
    if not args.workbook.is_file():
        parser.error(f"Workbook not found: {args.workbook}")
    payload = build_payload(args.workbook)
    args.output.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    included = sum(
        sheet["status"] in {"loaded", "included-placeholder"}
        for sheet in payload["sheetAudit"]
    )
    print(
        f"Wrote {len(payload['records'])} metric rows, "
        f"{len(payload['estore'])} eStore rows, {len(payload['campaigns'])} campaigns, "
        f"{len(payload['forms'])} forms, and "
        f"{sum(len(rows) for rows in payload['audienceDimensions'].values())} audience rows "
        f"from {included}/{len(payload['sheetAudit'])} "
        f"included workbook sheets to {args.output}."
    )
    for sheet in payload["sheetAudit"]:
        if sheet["status"] not in {"loaded", "included-placeholder"}:
            print(f"Skipped {sheet['name']}: {sheet['status']}")
        elif sheet["status"] == "included-placeholder":
            print(f"Included with placeholder warning: {sheet['name']}")


if __name__ == "__main__":
    main()
