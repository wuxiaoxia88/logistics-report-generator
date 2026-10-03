"""Regression tests for validated, atomic report delivery."""

import datetime as dt
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

from pypdf import PdfWriter

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import report_delivery as delivery


def pdf_bytes(pages=2):
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=595, height=842)
    stream = io.BytesIO()
    writer.write(stream)
    return stream.getvalue()


class DeliveryTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)
        self.html = self.root / "报告 with spaces #1.html"
        self.html.write_text("<!doctype html><title>合成测试</title>", encoding="utf-8")
        self.pdf = self.root / "report.pdf"

    def _render(self, content, code=0, stderr=""):
        def render(command, **kwargs):
            output = next(
                part for part in command if part.startswith("--print-to-pdf=")
            )
            Path(output.split("=", 1)[1]).write_bytes(content)
            return subprocess.CompletedProcess(command, code, "", stderr)

        return render

    def _assert_old_preserved(self):
        self.assertEqual(self.pdf.read_bytes(), b"old report")
        self.assertEqual(list(self.root.glob(".report-pdf-*")), [])

    def test_success_validates_new_pdf_and_returns_actual_page_count(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess, "run", side_effect=self._render(pdf_bytes(3))
        ):
            self.assertEqual(delivery.to_pdf("chrome", self.html, self.pdf), 3)
        self.assertEqual(delivery.count_pages(self.pdf), 3)
        self.assertEqual(list(self.root.glob(".report-pdf-*")), [])

    def test_nonzero_renderer_cannot_claim_stale_pdf_as_success(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess,
            "run",
            side_effect=self._render(pdf_bytes(), 1, "failed"),
        ):
            with self.assertRaisesRegex(RuntimeError, "exit code 1"):
                delivery.to_pdf("chrome", self.html, self.pdf)
        self._assert_old_preserved()

    def test_missing_renderer_output_does_not_reuse_stale_pdf(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess,
            "run",
            return_value=subprocess.CompletedProcess([], 0, "", ""),
        ):
            with self.assertRaisesRegex(RuntimeError, "no nonempty PDF"):
                delivery.to_pdf("chrome", self.html, self.pdf)
        self._assert_old_preserved()

    def test_empty_pdf_preserves_old_report(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(delivery.subprocess, "run", side_effect=self._render(b"")):
            with self.assertRaisesRegex(RuntimeError, "no nonempty PDF"):
                delivery.to_pdf("chrome", self.html, self.pdf)
        self._assert_old_preserved()

    def test_invalid_pdf_preserves_old_report(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess, "run", side_effect=self._render(b"%PDF-1.4\nbroken")
        ):
            with self.assertRaisesRegex(RuntimeError, "failed validation"):
                delivery.to_pdf("chrome", self.html, self.pdf)
        self._assert_old_preserved()

    def test_timeout_preserves_old_report_and_cleans_temporary_directory(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess,
            "run",
            side_effect=subprocess.TimeoutExpired("chrome", 120),
        ):
            with self.assertRaisesRegex(RuntimeError, "timed out"):
                delivery.to_pdf("chrome", self.html, self.pdf)
        self._assert_old_preserved()

    def test_missing_executable_has_clear_error(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess, "run", side_effect=FileNotFoundError("private/path")
        ):
            with self.assertRaisesRegex(RuntimeError, "Could not launch") as failure:
                delivery.to_pdf("chrome", self.html, self.pdf)
        self.assertNotIn("private/path", str(failure.exception))
        self._assert_old_preserved()

    def test_renderer_receives_encoded_file_uri_and_temporary_output(self):
        with patch.object(
            delivery.subprocess, "run", side_effect=self._render(pdf_bytes())
        ) as run:
            delivery.to_pdf("chrome", self.html, self.pdf)
        command = run.call_args.args[0]
        self.assertEqual(command[-1], self.html.resolve().as_uri())
        self.assertIn("%20", command[-1])
        self.assertIn("%23", command[-1])
        self.assertNotIn("--no-sandbox", command)
        target = Path(
            next(x for x in command if x.startswith("--print-to-pdf=")).split("=", 1)[1]
        )
        self.assertNotEqual(target, self.pdf)
        self.assertEqual(target.parent.parent, self.pdf.parent.resolve())

    def test_error_diagnostics_are_bounded_and_strip_control_characters(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess,
            "run",
            side_effect=self._render(b"", 7, "x" * 10000 + "\x00\nEND"),
        ):
            with self.assertRaises(RuntimeError) as failure:
                delivery.to_pdf("chrome", self.html, self.pdf)
        self.assertLess(len(str(failure.exception)), 1100)
        self.assertNotIn("\x00", str(failure.exception))
        self.assertNotIn("\n", str(failure.exception))
        self.assertTrue(str(failure.exception).endswith("END"))

    def test_failed_pdf_replace_preserves_old_report_and_cleans_up(self):
        self.pdf.write_bytes(b"old report")
        with patch.object(
            delivery.subprocess, "run", side_effect=self._render(pdf_bytes())
        ):
            with patch.object(delivery.os, "replace", side_effect=OSError("failed")):
                with self.assertRaisesRegex(RuntimeError, "Could not save"):
                    delivery.to_pdf("chrome", self.html, self.pdf)
        self._assert_old_preserved()

    def test_renderer_log_is_captured_without_inheritable_pipe_waits(self):
        self.pdf.write_bytes(b"old report")

        def render(command, **kwargs):
            self.assertEqual(kwargs["stdout"], subprocess.DEVNULL)
            self.assertNotIn("capture_output", kwargs)
            kwargs["stderr"].write(b"synthetic renderer failure")
            return subprocess.CompletedProcess(command, 2)

        with patch.object(delivery.subprocess, "run", side_effect=render):
            with self.assertRaisesRegex(RuntimeError, "synthetic renderer failure"):
                delivery.to_pdf("chrome", self.html, self.pdf)
        self._assert_old_preserved()

    def test_page_count_uses_parser(self):
        self.pdf.write_bytes(pdf_bytes(4))
        self.assertEqual(delivery.count_pages(self.pdf), 4)

    def test_page_count_rejects_missing_or_invalid_or_zero_page_files(self):
        with self.assertRaises(ValueError):
            delivery.count_pages(self.pdf)
        for content in (b"", b"not a PDF", b"%PDF-1.4\nbroken", pdf_bytes(0)):
            self.pdf.write_bytes(content)
            with self.assertRaises(ValueError):
                delivery.count_pages(self.pdf)

    def test_page_count_requires_pypdf(self):
        self.pdf.write_bytes(pdf_bytes())
        with patch.dict(sys.modules, {"pypdf": None}):
            with self.assertRaisesRegex(ValueError, "requires the pypdf"):
                delivery.count_pages(self.pdf)

    def test_find_chrome_explicit_precedes_environment(self):
        with patch.dict(os.environ, {"CHROME_BIN": "environment-browser"}):
            with patch.object(
                delivery, "_executable", return_value="explicit-browser"
            ) as executable:
                self.assertEqual(delivery.find_chrome("explicit"), "explicit-browser")
        executable.assert_called_once_with("explicit")

    def test_find_chrome_invalid_configuration_does_not_fallback(self):
        with patch.dict(os.environ, {"CHROME_BIN": "missing-private-path"}):
            with patch.object(delivery, "_executable", return_value=None) as executable:
                self.assertIsNone(delivery.find_chrome())
        executable.assert_called_once_with("missing-private-path")

    def test_find_chrome_checks_windows_and_path_candidates(self):
        seen = []

        def resolve(candidate):
            seen.append(candidate)
            return "/usr/bin/chromium" if candidate == "chromium" else None

        with patch.dict(os.environ, {"PROGRAMFILES": "C:/Programs"}, clear=True):
            with patch.object(delivery, "_executable", side_effect=resolve):
                self.assertEqual(delivery.find_chrome(), "/usr/bin/chromium")
        self.assertIn("C:/Programs/Google/Chrome/Application/chrome.exe", seen)
        self.assertIn(
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", seen
        )

    def test_executable_resolves_path_commands(self):
        with patch.object(
            delivery.shutil, "which", return_value="/usr/bin/chromium"
        ) as which:
            self.assertEqual(delivery._executable("chromium"), "/usr/bin/chromium")
        which.assert_called_once_with("chromium")

    def test_basename_contains_actual_period_and_sanitizes_client(self):
        name = delivery.artifact_basename(
            "  卡/乐:\n客户  ", "weekly", dt.date(2026, 10, 1), "2026-10-07"
        )
        self.assertEqual(name, "卡乐客户_物流周报_20261001-20261007")
        self.assertIn(
            "物流月报",
            delivery.artifact_basename("卡乐", "monthly", "2026-10-01", "2026-10-31"),
        )

    def test_basename_handles_empty_long_and_reserved_client_names(self):
        for client in ("", "\x00/::", "😀" * 1000, "CON"):
            name = delivery.artifact_basename(
                client, "weekly", "2026-10-01", "2026-10-07"
            )
            self.assertLess(len(name.encode("utf-8")), 255)
            self.assertTrue(name.split("_物流")[0])
            self.assertNotRegex(name, r'[\\/:*?"<>|\x00-\x1f\x7f-\x9f]')

    def test_basename_rejects_invalid_modes_and_periods(self):
        for mode, start, end in (
            ("bad", "2026-10-01", "2026-10-07"),
            ("weekly", "bad", "2026-10-07"),
            ("weekly", "2026-10-07", "2026-10-01"),
        ):
            with self.assertRaises(ValueError):
                delivery.artifact_basename("卡乐", mode, start, end)

    def test_json_atomic_write_and_strict_serialization(self):
        manifest = self.root / "nested" / "manifest.json"
        delivery.write_json_atomic(manifest, {"客户": "卡乐", "pages": 2})
        self.assertEqual(
            json.loads(manifest.read_text(encoding="utf-8")),
            {"客户": "卡乐", "pages": 2},
        )
        old = manifest.read_bytes()
        with self.assertRaises(ValueError):
            delivery.write_json_atomic(manifest, {"value": float("nan")})
        self.assertEqual(manifest.read_bytes(), old)
        self.assertEqual(list(manifest.parent.glob(".report-json-*")), [])

    def test_failed_json_replace_preserves_old_manifest_and_cleans_up(self):
        manifest = self.root / "manifest.json"
        manifest.write_text('{"old": true}', encoding="utf-8")
        with patch.object(delivery.os, "replace", side_effect=OSError("failed")):
            with self.assertRaises(OSError):
                delivery.write_json_atomic(manifest, {"new": True})
        self.assertEqual(json.loads(manifest.read_text()), {"old": True})
        self.assertEqual(list(self.root.glob(".report-json-*")), [])


if __name__ == "__main__":
    unittest.main()
