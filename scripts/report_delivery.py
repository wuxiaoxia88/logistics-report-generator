"""Validated, atomic delivery of report files.

Rendering a PDF is successful only after the new file can be parsed. Existing
reports are left intact whenever rendering or validation fails.
"""

from __future__ import annotations

import datetime as dt
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import tempfile
from typing import Any


PDF_TIMEOUT_SECONDS = 120
_MAX_STDERR_CHARS = 1000
_CHROME_NAMES = (
    "google-chrome",
    "google-chrome-stable",
    "chromium",
    "chromium-browser",
    "chrome",
    "msedge",
)
_MAC_CHROME_PATHS = (
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Chromium.app/Contents/MacOS/Chromium",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
)


def _executable(candidate: str) -> str | None:
    """Resolve a configured file or a command on PATH without launching it."""
    candidate = candidate.strip()
    if not candidate:
        return None
    path = Path(candidate).expanduser()
    if path.is_file() and os.access(path, os.X_OK):
        return str(path.resolve())
    return shutil.which(candidate)


def find_chrome(explicit: str | None = None) -> str | None:
    """Find Chrome/Chromium/Edge on macOS, Windows, or Linux.

    An explicit setting, or CHROME_BIN when no explicit setting is supplied,
    takes precedence. An invalid configured executable returns None so that
    a configuration mistake does not silently select a different browser.
    """
    configured = explicit if explicit is not None else os.environ.get("CHROME_BIN")
    if configured is not None:
        return _executable(configured)

    candidates = list(_MAC_CHROME_PATHS)
    for variable in ("PROGRAMFILES", "PROGRAMFILES(X86)", "LOCALAPPDATA"):
        base = os.environ.get(variable)
        if base:
            for suffix in (
                "Google/Chrome/Application/chrome.exe",
                "Chromium/Application/chrome.exe",
                "Microsoft/Edge/Application/msedge.exe",
            ):
                candidates.append(str(Path(base) / suffix))
    candidates.extend(_CHROME_NAMES)
    for candidate in candidates:
        found = _executable(candidate)
        if found:
            return found
    return None


def count_pages(path: str | os.PathLike[str]) -> int:
    """Read a nonempty PDF with pypdf; invalid/unreadable PDFs raise ValueError."""
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ValueError("PDF validation requires the pypdf package.") from None

    try:
        with Path(path).open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise ValueError("Invalid PDF header.")
            stream.seek(0)
            reader = PdfReader(stream, strict=True)
            if reader.is_encrypted:
                raise ValueError("Encrypted PDFs cannot be validated.")
            pages = len(reader.pages)
            if pages < 1:
                raise ValueError("PDF contains no pages.")
            return pages
    except (OSError, ValueError) as error:
        raise ValueError("PDF is unreadable or invalid: %s" % error) from None
    except Exception:
        # pypdf raises several parser-specific exceptions. Do not expose its
        # potentially extensive traceback or untrusted file content.
        raise ValueError("PDF could not be parsed.") from None


def _stderr_summary(stderr: str | bytes | None) -> str:
    if not stderr:
        return ""
    if isinstance(stderr, bytes):
        stderr = stderr.decode("utf-8", errors="replace")
    cleaned = re.sub(r"[\x00-\x1f\x7f-\x9f]", " ", stderr).strip()
    return cleaned[-_MAX_STDERR_CHARS:]


def to_pdf(
    chrome: str,
    html_path: str | os.PathLike[str],
    pdf_path: str | os.PathLike[str],
) -> int:
    """Render, validate, and atomically replace a PDF; return its page count.

    Raise RuntimeError on any rendering, validation, or replacement failure.
    A stale target file is never used as evidence of successful rendering.
    """
    source = Path(html_path).resolve()
    target = Path(pdf_path).resolve()
    if not source.is_file():
        raise RuntimeError("Report HTML file does not exist.")

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(
            prefix=".report-pdf-", dir=target.parent
        ) as scratch:
            temporary_pdf = Path(scratch) / "report.pdf"
            command = [
                chrome,
                "--headless",
                "--disable-gpu",
                "--no-pdf-header-footer",
                "--print-to-pdf=" + str(temporary_pdf),
                source.as_uri(),
            ]
            try:
                # Chrome may start updater helpers that inherit stdio. Using
                # PIPE can then wait for a helper after the renderer has exited.
                # A temporary log lets us wait only for the renderer process.
                with (Path(scratch) / "renderer.stderr").open("w+b") as diagnostics:
                    result = subprocess.run(
                        command,
                        stdout=subprocess.DEVNULL,
                        stderr=diagnostics,
                        timeout=PDF_TIMEOUT_SECONDS,
                        check=False,
                    )
                    diagnostics.seek(0, os.SEEK_END)
                    diagnostics.seek(max(0, diagnostics.tell() - 4096))
                    stderr_tail = diagnostics.read()
            except subprocess.TimeoutExpired:
                raise RuntimeError(
                    "PDF rendering timed out after %d seconds." % PDF_TIMEOUT_SECONDS
                ) from None
            except OSError:
                raise RuntimeError("Could not launch the PDF renderer.") from None

            if result.returncode != 0:
                summary = _stderr_summary(stderr_tail or result.stderr)
                detail = ": " + summary if summary else ""
                raise RuntimeError(
                    "PDF renderer failed with exit code %d%s"
                    % (result.returncode, detail)
                )
            if not temporary_pdf.is_file() or temporary_pdf.stat().st_size == 0:
                raise RuntimeError("PDF renderer produced no nonempty PDF file.")
            try:
                pages = count_pages(temporary_pdf)
            except ValueError as error:
                raise RuntimeError("New PDF failed validation: %s" % error) from None
            os.replace(temporary_pdf, target)
            return pages
    except RuntimeError:
        raise
    except OSError:
        raise RuntimeError("Could not save the validated PDF file.") from None


def _period_date(value: dt.date | str) -> dt.date:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        try:
            return dt.date.fromisoformat(value)
        except ValueError:
            pass
    raise ValueError("Report period dates must be date objects or ISO dates.")


def _safe_client(client: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f\x7f-\x9f]', "", str(client))
    cleaned = re.sub(r"\s+", "_", cleaned).strip(" ._")
    cleaned = cleaned[:64]
    # Bound UTF-8 bytes as well as character count for common filesystem limits.
    while len(cleaned.encode("utf-8")) > 120:
        cleaned = cleaned[:-1]
    cleaned = cleaned.rstrip(" ._") or "客户"
    if re.fullmatch(r"(?i)(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])", cleaned):
        cleaned = "_" + cleaned
    return cleaned


def artifact_basename(
    client: str,
    mode: str,
    period_start: dt.date | str,
    period_end: dt.date | str,
) -> str:
    """Create a portable report basename containing its actual period."""
    if mode not in ("weekly", "monthly"):
        raise ValueError("Report mode must be weekly or monthly.")
    start, end = _period_date(period_start), _period_date(period_end)
    if end < start:
        raise ValueError("Report period ends before it starts.")
    report_type = "月报" if mode == "monthly" else "周报"
    return "%s_物流%s_%s-%s" % (
        _safe_client(client),
        report_type,
        start.strftime("%Y%m%d"),
        end.strftime("%Y%m%d"),
    )


def write_json_atomic(path: str | os.PathLike[str], payload: Any) -> None:
    """Serialize strict JSON before atomically replacing the destination."""
    destination = Path(path).resolve()
    # Validate serialization first: NaN/Infinity are not JSON and must not enter
    # an audit manifest, nor damage an existing manifest.
    content = json.dumps(payload, ensure_ascii=False, indent=2, allow_nan=False) + "\n"
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            prefix=".report-json-",
            suffix=".tmp",
            dir=destination.parent,
            delete=False,
        ) as stream:
            temporary_path = Path(stream.name)
            stream.write(content)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary_path, destination)
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
