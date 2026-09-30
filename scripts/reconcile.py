"""Run the versioned reconciliation engine and save reproducible report data.

This entry point never overwrites inputs or sends business data anywhere.
Excel presentation must use the saved result, not recompute its business rules.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import platform
import sys
from zipfile import BadZipFile
from decimal import Decimal, localcontext
from pathlib import Path

import openpyxl

from reconciliation_core import reconcile

SKILL_ROOT = Path(__file__).resolve().parents[1]
DEPENDENCY_VERSION = "3.1.5"
MANIFEST_VERSION = "1"


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def serial(value):
    if isinstance(value, Decimal):
        if not value.is_finite():
            raise ValueError("结果中存在非有限数值")
        return format(value, "f")
    raise TypeError(f"结果包含不支持的类型: {type(value).__name__}")


def canonical_bytes(value) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, indent=2,
                       allow_nan=False, default=serial) + "\n").encode("utf-8")


def unique_keys(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"配置存在重复键: {key}")
        result[key] = value
    return result


def read_profile(path: Path) -> dict:
    result = json.loads(path.read_text(encoding="utf-8-sig"), object_pairs_hook=unique_keys,
                        parse_constant=lambda value: (_ for _ in ()).throw(ValueError(f"非法数值: {value}")))
    if not isinstance(result, dict):
        raise ValueError("规则配置必须为 JSON 对象")
    return result


def csv_value(value):
    if value is None:
        return ""
    if isinstance(value, Decimal):
        return serial(value)
    if isinstance(value, (list, dict)):
        value = json.dumps(value, ensure_ascii=False, sort_keys=True, default=serial)
    if isinstance(value, str) and value.lstrip().startswith(("=", "+", "-", "@", "\t", "\r")):
        # Protect text fields when a CSV is opened in spreadsheet software.
        return "'" + value
    return value


TABLES = {
    "products": ("商品异常汇总.csv", ["商品编码", "异常总单量", "总差额", "正差额", "负差绝对额", "原因单量", "首找对象", "处理要求"]),
    "details": ("运单明细.csv", ["商品编码", "物流单号", "主原因", "首找对象", "快递重量kg", "系统重量kg", "快递账单收费", "按系统重量计算", "差额", "处理要求"]),
    "evidence": ("快递多收费举证.csv", ["物流单号", "商品编码", "主原因", "快递账单收费", "按系统重量计算", "差额", "重量影响金额", "计价差异金额"]),
    "unidentified_records": ("待补单号记录.csv", ["来源", "源工作表", "源行", "原始单号", "原因"]),
    "rule_issues": ("价格表待核事项.csv", ["问题", "来源", "影响范围"]),
}


def write_tables(result: dict, output: Path) -> list[str]:
    names = []
    for key, (name, preferred) in TABLES.items():
        rows = result.get(key, [])
        columns = set().union(*(row.keys() for row in rows)) if rows else set(preferred)
        fields = [field for field in preferred if field in columns]
        fields += sorted(columns.difference(fields))
        with (output / name).open("w", encoding="utf-8-sig", newline="") as stream:
            writer = csv.DictWriter(stream, fieldnames=fields, lineterminator="\n")
            writer.writeheader()
            for row in rows:
                writer.writerow({field: csv_value(row.get(field)) for field in fields})
        names.append(name)
    return names


def validate_result(result: dict):
    required = {"engine_version", "summary", "details", "all_records", "products",
                "evidence", "price_rules", "rule_issues", "validation"}
    if not required <= result.keys():
        raise ValueError("引擎结果结构不完整")
    rows = result["details"]
    ids = [row["物流单号"] for row in rows]
    if len(ids) != len(set(ids)) or any(not tracking for tracking in ids):
        raise ValueError("异常运单单号不唯一或为空")
    if sum(row["异常总单量"] for row in result["products"]) != len(rows):
        raise ValueError("商品汇总与异常明细单量不守恒")
    expected_evidence = set()
    for row in result["all_records"]:
        a, c, s = (row.get(key) for key in ("快递账单收费", "按快递重量计算", "按系统重量计算"))
        total = row.get("差额")
        if a is not None and s is not None and total != a - s:
            raise ValueError("总差额与基础金额不一致")
        if a is not None and c is not None and s is not None:
            w, p = row.get("重量影响金额"), row.get("计价差异金额")
            if w != c - s or p != a - c or total != w + p:
                raise ValueError("差额分解不守恒")
            if total > 0 and row.get("主原因") in {"重量影响运费", "计价不符", "重量＋计价差异"}:
                expected_evidence.add(row["物流单号"])
    evidence_ids = [row["物流单号"] for row in result["evidence"]]
    if len(evidence_ids) != len(set(evidence_ids)) or set(evidence_ids) != expected_evidence:
        raise ValueError("多收费举证集合与明确正差运单不一致")
    for row in result["evidence"]:
        if row["差额"] is None or row["差额"] <= 0:
            raise ValueError("举证明细包含未知或非正差金额")
    if any(value is False for value in result["validation"].values()):
        raise ValueError("引擎校验存在失败项")


def run(args) -> dict:
    if sys.version_info < (3, 10):
        raise ValueError("需要 Python 3.10 或以上版本")
    if openpyxl.__version__ != DEPENDENCY_VERSION:
        raise ValueError("依赖版本不一致；请运行 python -m pip install -r requirements.txt")
    files = {role: Path(getattr(args, role)).resolve() for role in ("courier", "system", "price")}
    if len(set(files.values())) != 3:
        raise ValueError("三种输入必须分别指定，不能把同一文件同时当作两种账单")
    for path in files.values():
        if not path.is_file() or path.suffix.lower() != ".xlsx":
            raise ValueError(f"需要存在的 .xlsx 输入文件: {path.name}")
    output = Path(args.output_dir).resolve()
    if output.exists() and (not output.is_dir() or any(output.iterdir())):
        raise ValueError("输出目录非空；请指定一个新目录，以免混入旧结果")
    profile = read_profile(Path(args.profile))
    hashes = {role: file_sha256(path) for role, path in files.items()}
    result = reconcile(files["courier"], files["system"], files["price"], profile)
    with localcontext() as context:
        context.prec = 50
        validate_result(result)
    if hashes != {role: file_sha256(path) for role, path in files.items()}:
        raise ValueError("运行期间输入文件发生变化，结果未交付；请使用静态副本重跑")
    payload = canonical_bytes(result)
    result_hash = hashlib.sha256(payload).hexdigest()
    comparison = None
    if args.compare:
        previous = json.loads(Path(args.compare).read_text(encoding="utf-8"))
        previous_hash = hashlib.sha256(canonical_bytes(previous)).hexdigest()
        comparison = {"same_result": previous_hash == result_hash,
                      "expected_result_sha256": previous_hash}
        if not comparison["same_result"]:
            raise ValueError("复现比对未通过：本次计算结果与指定 result.json 不同。请检查输入、配置和引擎版本")
    output.mkdir(parents=True, exist_ok=True)
    (output / "result.json").write_bytes(payload)
    config_payload = canonical_bytes(profile)
    (output / "profile-resolved.json").write_bytes(config_payload)
    artifacts = ["result.json", "profile-resolved.json", *write_tables(result, output)]
    code_files = sorted((SKILL_ROOT / "scripts").glob("*.py"))
    code_hashes = {path.relative_to(SKILL_ROOT).as_posix(): file_sha256(path) for path in code_files}
    manifest = {
        "manifest_version": MANIFEST_VERSION,
        "engine_version": result["engine_version"],
        "profile_version": result.get("profile_version", profile.get("profile_version")),
        "config_revision": profile.get("config_revision"),
        "inputs": {role: {"filename": path.name, "sha256": hashes[role]} for role, path in files.items()},
        "profile_sha256": hashlib.sha256(config_payload).hexdigest(),
        "code_sha256": code_hashes,
        "result_sha256": result_hash,
        "environment": {"python": platform.python_version(), "openpyxl": openpyxl.__version__},
        "validation": result["validation"],
        "comparison": comparison,
        "artifacts": {name: file_sha256(output / name) for name in artifacts},
    }
    (output / "manifest.json").write_bytes(canonical_bytes(manifest))
    return {"engine_version": result["engine_version"], "result_sha256": result_hash,
            "summary": result["summary"], "output_dir": str(output)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--courier", required=True)
    parser.add_argument("--system", required=True)
    parser.add_argument("--price", required=True)
    parser.add_argument("--profile", default=str(SKILL_ROOT / "profiles" / "yuantong-jushuitan-v1.json"))
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--compare", help="Optional prior result.json; fail if the complete canonical result differs")
    args = parser.parse_args()
    try:
        print(canonical_bytes(run(args)).decode("utf-8"), end="")
        return 0
    except (ValueError, OSError, KeyError, TypeError, AssertionError, BadZipFile) as error:
        print(f"未完成核对：{error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
