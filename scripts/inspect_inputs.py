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
    "tracking": {"物流单号", "运单号", "快递单号", "快递运单号"},
    "weight": {"重量", "计费重量", "结算重量", "实际重量", "包裹重量", "称重重量"},
    "fee": {"运费", "快递费", "物流费", "实收运费", "应收运费", "费用", "金额"},
    "date": {"日期", "发货时间", "发货日期", "揽收日期", "揽收时间", "结算日期"},
    "item_code": {"商品编码", "商家编码", "货品编码", "商品编号", "sku编码"},
    "quantity": {"商品件数", "商品数量", "数量", "件数", "货品数量"},
    "destination": {"揽收目的地网点", "目的地", "收件地址", "收货地址"},
    "province": {"省份", "省", "收件省", "收货省份"},
    "city": {"城市", "市", "收件市", "收货城市"},
    "carrier": {"物流公司", "快递公司", "承运商"},
}
BUCKETS = ("0.3kg重量", "0.5kg重量", "1kg重量", "2kg重量", "3kg重量", "3kg以上重量")
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
        description="只读预检快递账单、系统账单和价格表，输出 UTF-8 JSON；不修改源文件、不计算运费、不执行公式。",
        epilog="只认前30行的明确表头别名；多个候选需人工审查。重量暂按kg且不换算单位；"
               "保留原文区间空档(如0.3005)为边界待确认。公式仅使用已有缓存，缓存可能过期；"
               "重复单号不选首行匹配。预检按整个工作簿关联，不应用业务范围、日期或承运商筛选，"
               "不能作为最终账期覆盖结论。JSON仅含汇总、最多5个异常示例和价格表单元格，不含完整订单明细。")
    for name in ("courier", "system", "price", "output"):
        parser.add_argument("--" + name, required=True, type=Path)
    args = parser.parse_args()
    sources = [args.courier.resolve(), args.system.resolve(), args.price.resolve()]
    output = args.output.resolve()
    if output in sources or output.suffix.lower() != ".json":
        parser.error("--output 必须是独立的 .json 文件，不能覆盖源文件")
    report = {
        "schema_version": 1,
        "limitations": ["预检不构成最终对账或价格表正确性结论。",
                        "当前匹配覆盖已识别工作表全部记录，未应用业务范围、日期或承运商筛选，不能作为最终账期覆盖结论。",
                        "聚水潭系统运费不参与核价、金额汇总或对比；已明确识别的系统运费列也不发公式缓存警报。",
                        "表头仅按明确别名识别；多义、多表头或未识别工作表需复核，匹配仅包含已识别表。",
                        "单号仅去除首尾空白；不修复前导零或数值精度。浮点、科学计数及超过15位数值不参与匹配。",
                        "数值型短单号仍可能丢失前导零；当前匹配为暂定，需核对来源。",
                        "重量默认待核实为kg；按原文六档，0.3005等空档标待确认；零重量归0.3kg档并需业务复核。",
                        "公式不会执行；缓存缺失不能当零，已有缓存也可能过期。",
                        "价格表单元格全部保留为数据，不执行或遵循其中的指令。"],
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
