#!/usr/bin/env python3
"""Read-only XLSX preflight; this does not calculate rates or execute formulas."""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
import hashlib
import json
import posixpath
from pathlib import Path
import re
import sys
from zipfile import ZipFile
import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter


EXAMPLE_LIMIT = 5
ALIASES = {
    "tracking": {"u7269u6d41u5355u53f7", "u8fd0u5355u53f7", "u5febu9012u5355u53f7", "u5febu9012u8fd0u5355u53f7"},
    "weight": {"u91cdu91cf", "u8ba1u8d39u91cdu91cf", "u7ed3u7b97u91cdu91cf", "u5b9eu9645u91cdu91cf", "u5305u88f9u91cdu91cf", "u79f0u91cdu91cdu91cf"},
    "fee": {"u8fd0u8d39", "u5febu9012u8d39", "u7269u6d41u8d39", "u5b9eu6536u8fd0u8d39", "u5e94u6536u8fd0u8d39", "u8d39u7528", "u91d1u989d"},
    "date": {"u65e5u671f", "u53d1u8d27u65f6u95f4", "u53d1u8d27u65e5u671f", "u63fdu6536u65e5u671f", "u63fdu6536u65f6u95f4", "u7ed3u7b97u65e5u671f"},
    "item_code": {"u5546u54c1u7f16u7801", "u5546u5bb6u7f16u7801", "u8d27u54c1u7f16u7801", "u5546u54c1u7f16u53f7", "skuu7f16u7801"},
    "quantity": {"u5546u54c1u4ef6u6570", "u5546u54c1u6570u91cf", "u6570u91cf", "u4ef6u6570", "u8d27u54c1u6570u91cf"},
    "destination": {"u63fdu6536u76eeu7684u5730u7f51u70b9", "u76eeu7684u5730", "u6536u4ef6u5730u5740", "u6536u8d27u5730u5740"},
    "province": {"u7701u4efd", "u7701", "u6536u4ef6u7701", "u6536u8d27u7701u4efd"},
    "city": {"u57ceu5e02", "u5e02", "u6536u4ef6u5e02", "u6536u8d27u57ceu5e02"},
    "carrier": {"u7269u6d41u516cu53f8", "u5febu9012u516cu53f8", "u627fu8fd0u5546"},
}
BUCKETS = ("0.3kgu91cdu91cf", "0.5kgu91cdu91cf", "1kgu91cdu91cf", "2kgu91cdu91cf", "3kgu91cdu91cf", "3kgu4ee5u4e0au91cdu91cf")
BOUNDS = [(Decimal(a), Decimal(b)) for a, b in
          [("0", ".3"), (".301", ".5"), (".501", "1"), ("1.001", "2"), ("2.001", "3")]]
SCI = re.compile(r"^[+-]?(?:\d+(?:\.\d*)?|\.\d+)[eE][+-]?\d+$")
NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def blank(value):
    return value is None or (isinstance(value, str) and not value.strip())


def serial(value):
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    return str(getattr(value, "text", value))


def digest(path):
    result = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            result.update(chunk)
    return result.hexdigest()


def normalize_header(value):
    return re.sub(r"\s+", "", str(value or "")).lower()


def identify_headers(ws):
    found = []
    for row_number, row in enumerate(ws.iter_rows(min_row=1, max_row=30), 1):
        mapping = defaultdict(list)
        for col, cell in enumerate(row, 1):
            label = normalize_header(cell.value)
            for field, aliases in ALIASES.items():
                # Only explicit label aliases are candidates; no positional fallback.
                if label in aliases:
                    mapping[field].append({"column": get_column_letter(col), "index": col,
                                           "label": serial(cell.value)})
        if mapping.get("tracking"):
            found.append({"row": row_number, "candidates": dict(mapping)})
    return found


def read_merges(path):
    """Read merge declarations without loading a complete editable workbook."""
    with ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels}
        result = {}
        for sheet in workbook.find(NS + "sheets"):
            target = targets[sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]]
            member = posixpath.normpath(target.lstrip("/") if target.startswith("/") else "xl/" + target)
            tree = ET.fromstring(archive.read(member))
            result[sheet.attrib["name"]] = [item.attrib["ref"] for item in tree.iter(NS + "mergeCell")]
        return result


def id_value(value, number_format=""):
    if blank(value):
        return None, "blank"
    if isinstance(value, bool):
        return None, "unsupported_type"
    if isinstance(value, int):
        if len(str(abs(value))) > 15:
            return None, "numeric_over_15_digits"
        if value < 0:
            return None, "negative_numeric"
        if re.search(r"[eE][+-]0", number_format or ""):
            return None, "scientific_number_format"
        return str(value), "numeric_identifier_check_leading_zeros"
    if isinstance(value, float):
        return None, "floating_numeric_identifier"
    if isinstance(value, str):
        value = value.strip()
        if SCI.fullmatch(value):
            return None, "scientific_text_identifier"
        return value, "text"
    return None, "unsupported_type"


def weight_value(value):
    if blank(value):
        return None, None, "blank", False
    try:
        if isinstance(value, bool):
            raise InvalidOperation
        amount = Decimal(str(value).strip())
        if not amount.is_finite():
            raise InvalidOperation
    except (InvalidOperation, ValueError, TypeError):
        return None, None, "non_numeric", False
    precision = amount.as_tuple().exponent < -3
    if amount < 0:
        return amount, None, "negative", precision
    for index, (low, high) in enumerate(BOUNDS):
        if low <= amount <= high:
            return amount, BUCKETS[index], "valid", precision
    if amount > 3:
        return amount, BUCKETS[-1], "valid", precision
    return amount, None, "boundary_confirmation_required", precision


def limited_append(collection, value):
    if len(collection) < EXAMPLE_LIMIT:
        collection.append(value)


def profile_book(path, role):
    formulas_book = openpyxl.load_workbook(path, read_only=True, data_only=False, keep_links=False)
    values_book = openpyxl.load_workbook(path, read_only=True, data_only=True, keep_links=False)
    profile = {"path": str(path), "sha256": digest(path), "role": role, "sheets": []}
    records = []
    merge_map = read_merges(path) if role == "price" else {}
    try:
        for ws in formulas_book:
            cached_ws = values_book[ws.title]
            sheet = {"name": ws.title, "declared_rows": ws.max_row, "declared_columns": ws.max_column,
                     "actual_nonempty_rows": 0, "formula_count": 0,
                     "ignored_system_fee_formula_count": 0,
                     "formula_missing_cache_count": 0, "formula_examples": [],
                     "formula_missing_cache_examples": []}
            headers = identify_headers(ws) if role != "price" else []
            sheet["header_candidates"] = headers
            chosen = headers[0] if len(headers) == 1 else None
            chosen = chosen if chosen and all(len(chosen["candidates"].get(key, [])) == 1
                                               for key in ("tracking", "weight")) else None
            sheet["profiling_status"] = "explicit_unique_header" if chosen else (
                "price_cells_only" if role == "price" else "header_requires_review")
            indices = {field: items[0]["index"] - 1 for field, items in chosen["candidates"].items()
                       if len(items) == 1} if chosen else {}
            # Ignore every explicitly labeled system fee column, even when field
            # mapping is ambiguous. Do not read fee values to resolve ambiguity.
            fee_header_rows = {}
            if role == "system":
                for header in headers:
                    for candidate in header["candidates"].get("fee", []):
                        column_index = candidate["index"] - 1
                        fee_header_rows[column_index] = min(
                            fee_header_rows.get(column_index, header["row"]), header["row"])
            stats, examples, bucket_counts, ids = Counter(), defaultdict(list), Counter(), Counter()
            sheet_records = []
            if role == "price":
                sheet["merged_ranges"] = merge_map.get(ws.title, [])
                sheet["nonempty_cells"] = []
            # Traverse complete source/cache streams together; formulas are never evaluated.
            for row_number, (row, cached) in enumerate(zip(ws.iter_rows(), cached_ws.iter_rows()), 1):
                if not any(not blank(cell.value) for cell in row):
                    continue
                sheet["actual_nonempty_rows"] += 1
                for column_index, (cell, cache) in enumerate(zip(row, cached)):
                    ignored_fee = (column_index in fee_header_rows
                                   and row_number > fee_header_rows[column_index])
                    if cell.data_type == "f" and ignored_fee:
                        sheet["ignored_system_fee_formula_count"] += 1
                    elif cell.data_type == "f":
                        sheet["formula_count"] += 1
                        limited_append(sheet["formula_examples"], {"cell": cell.coordinate,
                                       "formula": serial(cell.value), "cached": serial(cache.value)})
                        if cache.value is None:
                            sheet["formula_missing_cache_count"] += 1
                            limited_append(sheet["formula_missing_cache_examples"], cell.coordinate)
                    if role == "price" and not blank(cell.value):
                        item = {"cell": cell.coordinate, "value": serial(cell.value)}
                        if cell.data_type == "f":
                            item["cached"] = serial(cache.value)
                        sheet["nonempty_cells"].append(item)
                if not chosen or row_number <= chosen["row"]:
                    continue
                stats["actual_data_rows"] += 1

                def get(field):
                    return cached[indices[field]].value if field in indices else None

                id_cell = row[indices["tracking"]]
                key, id_status = id_value(get("tracking"), id_cell.number_format)
                stats["id_" + id_status] += 1
                if id_status not in ("text",):
                    limited_append(examples["id_" + id_status], {"cell": id_cell.coordinate,
                                   "value": serial(get("tracking"))})
                if key is not None:
                    ids[key] += 1
                weight, bucket, weight_status, precision = weight_value(get("weight"))
                stats["weight_" + weight_status] += 1
                if precision:
                    stats["weight_more_than_3_decimal_places"] += 1
                    limited_append(examples["weight_more_than_3_decimal_places"], {
                        "cell": row[indices["weight"]].coordinate, "value": serial(get("weight"))})
                if weight_status != "valid":
                    limited_append(examples["weight_" + weight_status], {
                        "cell": row[indices["weight"]].coordinate, "value": serial(get("weight"))})
                if bucket:
                    bucket_counts[bucket] += 1
                missing_item_code = "item_code" not in indices or blank(get("item_code"))
                missing_quantity = "quantity" not in indices or blank(get("quantity"))
                missing_goods = missing_item_code or missing_quantity
                if role == "system":
                    stats["goods_information_missing_rows"] += int(missing_goods)
                    stats["item_code_missing_rows"] += int(missing_item_code)
                    stats["quantity_missing_rows"] += int(missing_quantity)
                if key is not None:
                    sheet_records.append({"id": key, "weight": weight, "bucket": bucket,
                                          "weight_valid": weight_status == "valid", "goods_missing": missing_goods,
                                          "item_code_missing": missing_item_code, "quantity_missing": missing_quantity})
            if chosen:
                for key in ("actual_data_rows", "id_blank", "id_numeric_over_15_digits",
                            "id_floating_numeric_identifier", "id_scientific_text_identifier",
                            "id_scientific_number_format", "weight_blank", "weight_negative",
                            "weight_non_numeric", "weight_more_than_3_decimal_places",
                            "weight_boundary_confirmation_required"):
                    stats.setdefault(key, 0)
                sheet["statistics"] = dict(stats)
                sheet["weight_buckets"] = {key: bucket_counts[key] for key in BUCKETS}
                sheet["examples"] = dict(examples)
                sheet["identifiers"] = identifier_summary(ids)
                sheet["goods_headers_available"] = all(key in indices for key in ("item_code", "quantity"))
                records.extend(sheet_records)
            profile["sheets"].append(sheet)
    finally:
        formulas_book.close()
        values_book.close()
    if role != "price":
        profile["identifiers_across_profiled_sheets"] = identifier_summary(Counter(r["id"] for r in records))
    return profile, records


def identifier_summary(counts):
    duplicate = [(key, count) for key, count in counts.items() if count > 1]
    return {"usable_identifier_rows": sum(counts.values()), "distinct_identifiers": len(counts),
            "duplicate_identifiers": len(duplicate), "rows_with_duplicate_identifiers": sum(n for _, n in duplicate),
            "excess_duplicate_rows": sum(n - 1 for _, n in duplicate),
            "duplicate_examples": [{"id": key, "rows": n} for key, n in duplicate[:EXAMPLE_LIMIT]]}


def match_summary(courier, system):
    left, right = defaultdict(list), defaultdict(list)
    for row in courier:
        left[row["id"]].append(row)
    for row in system:
        right[row["id"]].append(row)
    common = left.keys() & right.keys()
    left_only, right_only = left.keys() - right.keys(), right.keys() - left.keys()
    unique = {key for key in common if len(left[key]) == len(right[key]) == 1}
    counts = Counter()
    for key in unique:
        a, b = left[key][0], right[key][0]
        if a["weight"] is None or b["weight"] is None or a["weight"] < 0 or b["weight"] < 0:
            counts["raw_weight_unavailable"] += 1
        else:
            counts["raw_weight_equal" if a["weight"] == b["weight"] else "raw_weight_different"] += 1
        if not a["bucket"] or not b["bucket"]:
            counts["weight_bucket_unavailable"] += 1
        elif a["bucket"] == b["bucket"]:
            counts["weight_bucket_equal"] += 1
        else:
            counts["weight_bucket_different"] += 1
            if b["goods_missing"]:
                counts["bucket_mismatch_goods_information_missing"] += 1
            if b["item_code_missing"]:
                counts["bucket_mismatch_item_code_missing"] += 1
            if b["quantity_missing"]:
                counts["bucket_mismatch_quantity_missing"] += 1
        if b["goods_missing"]:
            counts["goods_information_missing"] += 1
        if b["item_code_missing"]:
            counts["item_code_missing"] += 1
        if b["quantity_missing"]:
            counts["quantity_missing"] += 1
    for name in ("raw_weight_unavailable", "raw_weight_equal", "raw_weight_different", "weight_bucket_unavailable",
                 "weight_bucket_equal", "weight_bucket_different", "goods_information_missing",
                 "bucket_mismatch_goods_information_missing", "item_code_missing", "quantity_missing",
                 "bucket_mismatch_item_code_missing", "bucket_mismatch_quantity_missing"):
        counts.setdefault(name, 0)
    return {"common_distinct_identifiers": len(common), "unique_one_to_one_pairs": len(unique),
            "common_identifiers_excluded_due_to_duplicates": len(common - unique),
            "courier_only_distinct_identifiers": len(left_only), "courier_only_rows": sum(len(left[k]) for k in left_only),
            "system_only_distinct_identifiers": len(right_only), "system_only_rows": sum(len(right[k]) for k in right_only),
            "courier_only_examples": sorted(left_only)[:EXAMPLE_LIMIT],
            "system_only_examples": sorted(right_only)[:EXAMPLE_LIMIT], "unique_pair_weight_checks": dict(counts)}


def main():
    parser = argparse.ArgumentParser(
        description="u53eau8bfbu9884u68c0u5febu9012u8d26u5355u3001u7cfbu7edfu8d26u5355u548cu4ef7u683cu8868uff0cu8f93u51fa UTF-8 JSONuff1bu4e0du4feeu6539u6e90u6587u4ef6u3001u4e0du8ba1u7b97u8fd0u8d39u3001u4e0du6267u884cu516cu5f0fu3002",
        epilog="u53eau8ba4u524d30u884cu7684u660eu786eu8868u5934u522bu540duff1bu591au4e2au5019u9009u9700u4ebau5de5u5ba1u67e5u3002u91cdu91cfu6682u6309kgu4e14u4e0du6362u7b97u5355u4f4duff1b"
               "u4fddu7559u539fu6587u533au95f4u7a7au6863(u59820.3005)u4e3au8fb9u754cu5f85u786eu8ba4u3002u516cu5f0fu4ec5u4f7fu7528u5df2u6709u7f13u5b58uff0cu7f13u5b58u53efu80fdu8fc7u671fuff1b"
               "u91cdu590du5355u53f7u4e0du9009u9996u884cu5339u914du3002u9884u68c0u6309u6574u4e2au5de5u4f5cu7c3fu5173u8054uff0cu4e0du5e94u7528u4e1au52a1u8303u56f4u3001u65e5u671fu6216u627fu8fd0u5546u7b5bu9009uff0c"
               "u4e0du80fdu4f5cu4e3au6700u7ec8u8d26u671fu8986u76d6u7ed3u8bbau3002JSONu4ec5u542bu6c47u603bu3001u6700u591a5u4e2au5f02u5e38u793au4f8bu548cu4ef7u683cu8868u5355u5143u683cuff0cu4e0du542bu5b8cu6574u8ba2u5355u660eu7ec6u3002")
    for name in ("courier", "system", "price", "output"):
        parser.add_argument("--" + name, required=True, type=Path)
    args = parser.parse_args()
    sources = [args.courier.resolve(), args.system.resolve(), args.price.resolve()]
    output = args.output.resolve()
    if output in sources or output.suffix.lower() != ".json":
        parser.error("--output u5fc5u987bu662fu72ecu7acbu7684 .json u6587u4ef6uff0cu4e0du80fdu8986u76d6u6e90u6587u4ef6")
    report = {
        "schema_version": 1,
        "limitations": ["u9884u68c0u4e0du6784u6210u6700u7ec8u5bf9u8d26u6216u4ef7u683cu8868u6b63u786eu6027u7ed3u8bbau3002",
                        "u5f53u524du5339u914du8986u76d6u5df2u8bc6u522bu5de5u4f5cu8868u5168u90e8u8bb0u5f55uff0cu672au5e94u7528u4e1au52a1u8303u56f4u3001u65e5u671fu6216u627fu8fd0u5546u7b5bu9009uff0cu4e0du80fdu4f5cu4e3au6700u7ec8u8d26u671fu8986u76d6u7ed3u8bbau3002",
                        "u805au6c34u6f6du7cfbu7edfu8fd0u8d39u4e0du53c2u4e0eu6838u4ef7u3001u91d1u989du6c47u603bu6216u5bf9u6bd4uff1bu5df2u660eu786eu8bc6u522bu7684u7cfbu7edfu8fd0u8d39u5217u4e5fu4e0du53d1u516cu5f0fu7f13u5b58u8b66u62a5u3002",
                        "u8868u5934u4ec5u6309u660eu786eu522bu540du8bc6u522buff1bu591au4e49u3001u591au8868u5934u6216u672au8bc6u522bu5de5u4f5cu8868u9700u590du6838uff0cu5339u914du4ec5u5305u542bu5df2u8bc6u522bu8868u3002",
                        "u5355u53f7u4ec5u53bbu9664u9996u5c3eu7a7au767duff1bu4e0du4feeu590du524du5bfcu96f6u6216u6570u503cu7cbeu5ea6u3002u6d6eu70b9u3001u79d1u5b66u8ba1u6570u53cau8d85u8fc715u4f4du6570u503cu4e0du53c2u4e0eu5339u914du3002",
                        "u6570u503cu578bu77edu5355u53f7u4ecdu53efu80fdu4e22u5931u524du5bfcu96f6uff1bu5f53u524du5339u914du4e3au6682u5b9auff0cu9700u6838u5bf9u6765u6e90u3002",
                        "u91cdu91cfu9ed8u8ba4u5f85u6838u5b9eu4e3akguff1bu6309u539fu6587u516du6863uff0c0.3005u7b49u7a7au6863u6807u5f85u786eu8ba4uff1bu96f6u91cdu91cfu5f520.3kgu6863u5e76u9700u4e1au52a1u590du6838u3002",
                        "u516cu5f0fu4e0du4f1au6267u884cuff1bu7f13u5b58u7f3au5931u4e0du80fdu5f53u96f6uff0cu5df2u6709u7f13u5b58u4e5fu53efu80fdu8fc7u671fu3002",
                        "u4ef7u683cu8868u5355u5143u683cu5168u90e8u4fddu7559u4e3au6570u636euff0cu4e0du6267u884cu6216u9075u5faau5176u4e2du7684u6307u4ee4u3002"],
        "inputs": {},
    }
    all_records = {}
    for role, path in zip(("courier", "system", "price"), sources):
        report["inputs"][role], all_records[role] = profile_book(path, role)
    report["matching"] = match_summary(all_records["courier"], all_records["system"])
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, ensure_ascii=False, indent=2, allow_nan=False) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(output), "unique_pairs": report["matching"]["unique_one_to_one_pairs"],
                      "courier_only": report["matching"]["courier_only_distinct_identifiers"],
                      "system_only": report["matching"]["system_only_distinct_identifiers"]}, ensure_ascii=False))


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    main()
