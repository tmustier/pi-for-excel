"""Shared LibreOffice headless recalc helper for eval fixtures.

Fidelity contract: LO recalc was validated against the Excel-derived
expected-clean.json for the gpu-farm model (35/35 at rel_tol 1e-6,
2026-07-08; see fixtures/gpu-farm/oracle_lo.py --validate-against).
Trusted for the vanilla function set (CHOOSE/IF/ISERROR/IFERROR/SUM/
SUMIF/SUMPRODUCT/MIN/MAX/AVERAGE). Re-validate before relying on it for
fixtures using dates, IRR/NPV, lookups, or text functions.

Uses an isolated LO user profile per call so concurrent soffice
instances (e.g. SpreadsheetBench harness runs) are undisturbed.
"""

import shutil
import subprocess
import tempfile
from pathlib import Path

SOFFICE = "/opt/homebrew/bin/soffice"


def lo_recalc(workbook: Path, out_path: Path, timeout: int = 180) -> Path:
    """Recalculate `workbook` via LO convert and write result to `out_path`.

    The input should have dirty formula cells (openpyxl round-trips drop
    cached values, which marks every formula dirty), so LO computes fresh
    values regardless of its recalc-on-load setting.
    """
    workbook = Path(workbook)
    out_path = Path(out_path)
    with tempfile.TemporaryDirectory() as td:
        tdp = Path(td)
        work = tdp / workbook.name
        shutil.copy2(workbook, work)
        proc = subprocess.run(
            [SOFFICE, "--headless", "--calc",
             f"-env:UserInstallation=file://{tdp / 'lo-profile'}",
             "--convert-to", "xlsx:Calc MS Excel 2007 XML",
             "--outdir", str(tdp / "out"), str(work)],
            capture_output=True, text=True, timeout=timeout,
        )
        conv = tdp / "out" / work.name
        if proc.returncode != 0 or not conv.exists():
            raise RuntimeError(
                f"LO recalc failed rc={proc.returncode}: {proc.stderr[-800:]}")
        out_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(conv, out_path)
    return out_path
