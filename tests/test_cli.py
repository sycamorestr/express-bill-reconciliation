"""End-to-end reproducibility and stale-output protection for the public command."""
from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from test_reconciliation import (
    ROOT, COURIER_HEADERS, SYSTEM_HEADERS,
    courier_row, system_row, write_table, write_prices,
)


class CommandTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory(prefix="reconciliation-command-")
        self.addCleanup(temporary.cleanup)
        self.directory = Path(temporary.name)
        self.courier = self.directory / "courier.xlsx"
        self.system = self.directory / "system.xlsx"
        self.price = self.directory / "price.xlsx"
        self.make_courier()
        write_table(self.system, "系统明细", SYSTEM_HEADERS,
                    [system_row("FAKE-CMD-OVER", .2), system_row("FAKE-CMD-UNDER", .2)])
        write_prices(self.price)

    def make_courier(self, charge=3):
        write_table(self.courier, "快递明细", COURIER_HEADERS,
                    [courier_row("FAKE-CMD-OVER", .4, charge), courier_row("FAKE-CMD-UNDER", .2, 1)])

    def command(self, output, *extra):
        return subprocess.run(
            [sys.executable, "-X", "utf8", str(ROOT / "scripts" / "reconcile.py"),
             "--courier", str(self.courier), "--system", str(self.system),
             "--price", str(self.price), "--output-dir", str(output), *map(str, extra)],
            capture_output=True, text=True, encoding="utf-8", timeout=45,
        )

    def test_two_runs_match_byte_for_byte_and_preserve_input_fingerprints(self):
        first, second = self.directory / "first", self.directory / "second"
        before = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in [self.courier, self.system, self.price]}
        run = self.command(first)
        self.assertEqual(run.returncode, 0, run.stderr)
        replay = self.command(second, "--compare", first / "result.json")
        self.assertEqual(replay.returncode, 0, replay.stderr)
        self.assertEqual((first / "result.json").read_bytes(), (second / "result.json").read_bytes())
        result = json.loads((first / "result.json").read_text(encoding="utf-8"))
        self.assertEqual(len(result["evidence"]), 1)
        self.assertEqual(result["evidence"][0]["差额"], "1.00")
        manifest = json.loads((second / "manifest.json").read_text(encoding="utf-8"))
        self.assertTrue(manifest["comparison"]["same_result"])
        for source in manifest["inputs"].values():
            self.assertEqual(source["sha256"], before[source["filename"]])
        for name, digest in manifest["artifacts"].items():
            self.assertEqual(hashlib.sha256((second / name).read_bytes()).hexdigest(), digest)

    def test_changed_inputs_fail_compare_without_publishing_output(self):
        first, second = self.directory / "first", self.directory / "second"
        run = self.command(first)
        self.assertEqual(run.returncode, 0, run.stderr)
        self.make_courier(charge=4)
        replay = self.command(second, "--compare", first / "result.json")
        self.assertEqual(replay.returncode, 2)
        self.assertIn("复现比对未通过", replay.stderr)
        self.assertFalse(second.exists())

    def test_existing_output_is_not_overwritten_or_mixed_with_new_results(self):
        output = self.directory / "existing"
        output.mkdir()
        sentinel = output / "previous.txt"
        sentinel.write_text("keep existing results", encoding="utf-8")
        run = self.command(output)
        self.assertEqual(run.returncode, 2)
        self.assertEqual(sentinel.read_text(encoding="utf-8"), "keep existing results")
        self.assertEqual(list(output.iterdir()), [sentinel])

    def test_supported_calendar_text_has_the_same_meaning_across_python_versions(self):
        from reconciliation_core import default_profile, reconcile
        for date in ["2040-09-10", "2040/9/10", "2040-09-10T00:00:00Z", "2040-09-10 08:00+08:00"]:
            with self.subTest(date=date):
                write_table(self.courier, "快递明细", COURIER_HEADERS,
                            [courier_row("FAKE-CMD-OVER", .4, 3, date=date)])
                result = reconcile(self.courier, self.system, self.price, default_profile())
                record = next(row for row in result["all_records"] if row["物流单号"] == "FAKE-CMD-OVER")
                self.assertEqual(record["快递日期"], "2040-09-10")
                self.assertEqual(str(record["差额"]), "1.00")
        for date in ["20400910", "2040-W37-1", "2040-09-10 99:99:99"]:
            with self.subTest(date=date):
                write_table(self.courier, "快递明细", COURIER_HEADERS,
                            [courier_row("FAKE-CMD-OVER", .4, 3, date=date)])
                result = reconcile(self.courier, self.system, self.price, default_profile())
                record = next(row for row in result["all_records"] if row["物流单号"] == "FAKE-CMD-OVER")
                self.assertIsNone(record["快递日期"])
                self.assertIsNone(record["差额"])
                self.assertEqual(record["主原因"], "日期待确认")

    def test_combination_quantity_text_is_preserved_as_source_evidence(self):
        from reconciliation_core import default_profile, reconcile
        combined = system_row("FAKE-CMD-OVER", .2, sku="虚构甲*1,虚构乙*2")
        combined["商品件数"] = "1,2"
        write_table(self.system, "系统明细", SYSTEM_HEADERS, [combined])
        result = reconcile(self.courier, self.system, self.price, default_profile())
        record = next(row for row in result["all_records"] if row["物流单号"] == "FAKE-CMD-OVER")
        self.assertEqual(record["商品件数"], "1,2")
        self.assertEqual(record["商品编码"], "虚构甲*1,虚构乙*2")
        self.assertNotIn("商品件数缺失", record["商品资料缺失"])
        self.assertEqual(str(record["差额"]), "1.00")


if __name__ == "__main__":
    unittest.main()
