"""User-visible report and CLI regressions; no browser is required."""

import contextlib
import io
import json
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import build_report as report
from report_core import analyze, load_many


class ReportRegression(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.current = self.root / "current.xlsx"
        self.output = self.root / "reports"

    def excel(self, rows, path=None):
        default = {
            "运单号": "0001",
            "寄件时间": pd.Timestamp("2026-09-14 09:00"),
            "签收时间": pd.Timestamp("2026-09-15 09:00"),
            "目的省份": "江苏",
            "件数": 1,
            "实际重量": 2,
            "结算重量": 2,
            "体积": 0.1,
            "备注": "",
        }
        pd.DataFrame([dict(default, **r) for r in rows]).to_excel(
            path or self.current, index=False
        )

    def run_cli(self, *extra, skip=True):
        args = [
            "--current",
            str(self.current),
            "--client",
            "测试客户",
            "--as-of",
            "2026-09-25T12:00:00+08:00",
            "--output",
            str(self.output),
        ]
        if skip:
            args.append("--skip-pdf")
        with (
            contextlib.redirect_stdout(io.StringIO()) as out,
            contextlib.redirect_stderr(io.StringIO()) as err,
        ):
            code = report.main(args + list(extra))
        return code, out.getvalue(), err.getvalue()

    def latest(self):
        return sorted(self.output.iterdir())[-1]

    def test_cli_emits_consistent_snapshot_and_quality(self):
        self.excel([{}, {"运单号": "0002", "备注": "在途丢失", "签收时间": pd.NaT}])
        code, out, err = self.run_cli()
        self.assertEqual(code, 0, err)
        manifest = json.loads((self.latest() / "manifest.json").read_text())
        snapshot = json.loads((self.latest() / "metrics.json").read_text())
        self.assertEqual(manifest["status"], "html_only")
        self.assertEqual(snapshot["metrics"]["运单数"], 2)
        self.assertEqual(snapshot["metrics"]["物流异常单数"], 1)
        self.assertEqual(snapshot["metrics"]["sla_rate"], 50)
        self.assertEqual(json.loads(out)["status"], "html_only")
        self.assertEqual(
            json.loads((self.latest() / "quality.json").read_text()),
            snapshot["metrics"]["quality"],
        )

    def test_bad_required_input_fails_without_success_report(self):
        pd.DataFrame({"运单号": ["0001"]}).to_excel(self.current, index=False)
        code, out, err = self.run_cli()
        self.assertEqual(code, 1)
        self.assertFalse(out)
        self.assertIn("必需列", err)
        self.assertFalse(self.output.exists())

    def test_empty_quality_csv_has_readable_schema(self):
        self.excel([{}])
        self.assertEqual(self.run_cli()[0], 0)
        data = pd.read_csv(self.latest() / "data_quality.csv")
        self.assertEqual(
            list(data.columns), ["file", "row", "field", "reason", "value"]
        )
        self.assertTrue(data.empty)

    def test_missing_optional_columns_degrade_without_false_failure_rates(self):
        pd.DataFrame(
            {"运单号": ["0001"], "寄件时间": [pd.Timestamp("2026-09-14 09:00")]}
        ).to_excel(self.current, index=False)
        code, _, err = self.run_cli()
        self.assertEqual(code, 0, err)
        saved = json.loads((self.latest() / "metrics.json").read_text())
        self.assertIsNone(saved["metrics"]["签收率"])
        self.assertIsNone(saved["metrics"]["物流异常率"])
        self.assertEqual(saved["metrics"]["sla_eligible"], 0)
        self.assertTrue(pd.read_csv(self.latest() / "province.csv").empty)
        report.validate_snapshot(saved)

    def test_incomplete_remark_coverage_does_not_recreate_zero_rate(self):
        pd.DataFrame(
            {"运单号": ["0001"], "寄件时间": [pd.Timestamp("2026-09-14 09:00")]}
        ).to_excel(self.current, index=False)
        stats = analyze(
            load_many([str(self.current)]), as_of="2026-09-25T12:00:00+08:00"
        )
        _, weekly = report.exception_daily(stats, "weekly")
        _, monthly = report.exception_daily(stats, "monthly")
        self.assertNotIn("0.0%", weekly)
        self.assertNotIn("0.0%", monthly)

    def test_manifest_artifact_hashes_match_outputs(self):
        import hashlib

        self.excel([{}])
        self.assertEqual(self.run_cli()[0], 0)
        manifest = json.loads((self.latest() / "manifest.json").read_text())
        self.assertIn("manifest.json", manifest["files"])
        for item in manifest["artifacts"]:
            payload = (self.latest() / item["path"]).read_bytes()
            self.assertEqual(item["bytes"], len(payload))
            self.assertEqual(item["sha256"], hashlib.sha256(payload).hexdigest())

    def test_low_sla_province_is_not_hidden_by_volume_cutoff(self):
        rows = []
        for province in range(9):
            for i in range(3):
                rows.append(
                    {
                        "运单号": f"P{province}-{i}",
                        "目的省份": f"省{province}",
                        "签收时间": pd.NaT
                        if province == 8
                        else pd.Timestamp("2026-09-15 09:00"),
                    }
                )
        self.excel(rows)
        stats = analyze(
            load_many([str(self.current)]), as_of="2026-09-25T12:00:00+08:00"
        )
        self.assertIn("省8", report.province(stats, 98))

    def test_future_rows_and_signs_are_in_quality_attachment(self):
        self.excel(
            [
                {"签收时间": pd.Timestamp("2026-09-26 12:00")},
                {
                    "运单号": "0002",
                    "寄件时间": pd.Timestamp("2026-09-20 18:00"),
                    "签收时间": pd.NaT,
                },
            ]
        )
        code, _, err = self.run_cli("--as-of", "2026-09-19T12:00:00+08:00")
        self.assertEqual(code, 0, err)
        quality = json.loads((self.latest() / "quality.json").read_text())
        self.assertEqual(quality["as_of_excluded_rows"], 1)
        self.assertEqual(quality["future_sign_rows"], 1)
        diagnostic = (self.latest() / "data_quality.csv").read_text(
            encoding="utf-8-sig"
        )
        self.assertIn("寄件晚于", diagnostic)
        self.assertIn("签收晚于", diagnostic)

    def test_report_text_is_escaped_and_promises_not_invented(self):
        self.excel(
            [{"备注": "在途丢失 <b id='remark'>标签</b>", "目的省份": "<b>江苏</b>"}]
        )
        code, _, err = self.run_cli("--client", "<b id='client'>客户</b>")
        self.assertEqual(code, 0, err)
        content = next(self.latest().glob("*.html")).read_text()
        self.assertNotIn("<b id='client'>", content)
        self.assertNotIn("<b id='remark'>", content)
        self.assertIn("&lt;b&gt;江苏&lt;/b&gt;", content)
        self.assertNotIn("2小时内响应", content)
        self.assertNotIn("时效达成良好", content)

    def test_no_signed_samples_does_not_display_100_percent_timing(self):
        self.excel([{"签收时间": pd.NaT}])
        stats = analyze(
            load_many([str(self.current)]), as_of="2026-09-15T12:00:00+08:00"
        )
        rates, table, _ = report.timing(stats)
        self.assertEqual(rates, [None, None, None, None])
        self.assertNotIn("100%", table)
        self.assertIn("待观察", report.province(stats, 98))

    def test_monthly_rows_weight_actual_shipments_and_peak(self):
        rows = [
            {
                "寄件时间": pd.Timestamp("2026-09-01 09:00"),
                "签收时间": pd.Timestamp("2026-09-05 13:00"),
            }
        ]
        rows += [
            {
                "运单号": f"A{i}",
                "寄件时间": pd.Timestamp("2026-09-02 09:00"),
                "签收时间": pd.Timestamp("2026-09-02 19:00"),
            }
            for i in range(9)
        ]
        rows += [
            {
                "运单号": f"B{i}",
                "寄件时间": pd.Timestamp("2026-09-08 09:00") + pd.Timedelta(days=i % 6),
                "签收时间": pd.Timestamp("2026-09-08 19:00") + pd.Timedelta(days=i % 6),
            }
            for i in range(42)
        ]
        self.excel(rows)
        data = load_many([str(self.current)])
        stats = analyze(data, as_of="2026-10-05T12:00:00+08:00")
        _, content = report.daily_timing(stats, "monthly", stats["_data"])
        self.assertIn("<td>第1周</td><td>10</td><td>10</td><td>19.0h</td>", content)
        self.assertIn("发货高峰第2周（42票）", report.overview(stats, "monthly", data))

    def test_actions_closed_requires_evidence_and_tracks_absence(self):
        path = self.root / "actions.json"
        path.write_text(
            json.dumps(
                [
                    {"waybill": "0001", "status": "closed"},
                    {"waybill": "0002", "status": "closed", "evidence": "记录"},
                ]
            )
        )
        actions = report.load_actions(path)
        self.assertFalse(actions[0]["verified_closed"])
        self.assertTrue(actions[1]["verified_closed"])
        summary = report.action_summary(
            actions,
            [{"waybill": "0003", "status": "in_progress"}],
            pd.Timestamp("2026-09-25T12:00:00+08:00"),
            "Asia/Shanghai",
        )
        self.assertIn("有闭环证据1票", summary)
        self.assertIn("1票未提供本期更新", summary)

    def test_actions_outside_client_shipments_fail(self):
        self.excel([{}])
        actions = self.root / "actions.json"
        actions.write_text('[{"waybill":"OTHER","status":"pending"}]')
        code, _, err = self.run_cli("--actions", str(actions))
        self.assertEqual(code, 1)
        self.assertIn("以外的运单", err)

    def test_history_snapshot_requires_matching_observation_lag(self):
        self.excel([{}])
        self.assertEqual(self.run_cli()[0], 0)
        saved = self.latest() / "metrics.json"
        self.excel(
            [
                {
                    "寄件时间": pd.Timestamp("2026-09-21 09:00"),
                    "签收时间": pd.Timestamp("2026-09-22 09:00"),
                }
            ]
        )
        code, _, err = self.run_cli(
            "--previous-report", str(saved), "--as-of", "2026-10-02T12:00:00+08:00"
        )
        self.assertEqual(code, 0, err)
        code, _, err = self.run_cli("--previous-report", str(saved))
        self.assertEqual(code, 1)
        self.assertIn("观察滞后不同", err)

    def test_history_rejects_malformed_nonfinite_and_inconsistent_metrics(self):
        self.excel([{}])
        self.assertEqual(self.run_cli()[0], 0)
        snapshot = json.loads((self.latest() / "metrics.json").read_text())
        self.excel(
            [
                {
                    "寄件时间": pd.Timestamp("2026-09-21 09:00"),
                    "签收时间": pd.Timestamp("2026-09-22 09:00"),
                }
            ]
        )
        bad = self.root / "bad.json"
        cases = [[], {**snapshot, "metrics": {}}]
        incorrect = json.loads(json.dumps(snapshot))
        incorrect["metrics"]["sla_rate"] = 1
        cases.append(incorrect)
        invalid = json.loads(json.dumps(snapshot))
        invalid["metrics"]["运单数"] = float("nan")
        cases.append(invalid)
        for case in cases:
            bad.write_text(json.dumps(case))
            code, _, err = self.run_cli(
                "--previous-report", str(bad), "--as-of", "2026-10-02T12:00:00+08:00"
            )
            self.assertEqual(code, 1, err)

    def test_excel_comparison_preserves_source_and_baseline(self):
        self.excel([{}])
        previous = self.root / "previous.xlsx"
        self.excel(
            [
                {
                    "寄件时间": pd.Timestamp("2026-09-07 09:00"),
                    "签收时间": pd.Timestamp("2026-09-08 09:00"),
                }
            ],
            previous,
        )
        code, _, err = self.run_cli("--previous", str(previous))
        self.assertEqual(code, 0, err)
        saved = json.loads((self.latest() / "metrics.json").read_text())
        baseline = saved["comparison"]
        self.assertEqual(baseline["source"]["kind"], "excel")
        self.assertEqual(baseline["metrics"]["运单数"], 1)
        self.assertEqual(len(baseline["source"]["input_files"][0]["hash"]), 64)

    def test_pdf_overflow_has_review_status_and_nonzero_exit(self):
        self.excel([{}])
        with (
            patch.object(report, "find_chrome", return_value="chrome"),
            patch.object(report, "to_pdf", return_value=3),
        ):
            code, _, err = self.run_cli(skip=False)
        self.assertEqual(code, 2)
        self.assertIn("两页", err)
        self.assertEqual(
            json.loads((self.latest() / "manifest.json").read_text())["status"],
            "needs_review",
        )

    def test_pdf_failure_has_failed_manifest(self):
        self.excel([{}])
        with (
            patch.object(report, "find_chrome", return_value="chrome"),
            patch.object(
                report, "to_pdf", side_effect=RuntimeError("conversion failed")
            ),
        ):
            code, _, _ = self.run_cli(skip=False)
        self.assertEqual(code, 1)
        self.assertEqual(
            json.loads((self.latest() / "manifest.json").read_text())["status"],
            "failed",
        )

    def test_repeated_runs_preserve_previous_artifacts(self):
        self.excel([{}])
        self.assertEqual(self.run_cli()[0], 0)
        first = self.latest()
        original = (first / "metrics.json").read_bytes()
        self.assertEqual(self.run_cli()[0], 0)
        self.assertEqual(len(list(self.output.iterdir())), 2)
        self.assertEqual((first / "metrics.json").read_bytes(), original)

    def test_csv_preserves_full_notes_and_neutralizes_formulas(self):
        content = "=SUM(1,2)" + "完整说明" * 100
        path = self.root / "notes.csv"
        report.csv_output(path, [{"note": content}])
        self.assertEqual(pd.read_csv(path).loc[0, "note"], "'" + content)

    def test_invalid_period_is_rejected(self):
        self.excel([{}])
        code, _, err = self.run_cli(
            "--period-start", "2026-09-01", "--period-end", "2026-09-20"
        )
        self.assertEqual(code, 1)
        self.assertIn("最多7天", err)


if __name__ == "__main__":
    unittest.main()
