"""Shared leakage/metadata scrub helpers for eval fixture builders.

Builders must call assert_no_leakage() on every emitted seed — it enforces
the full sweep from docs/proposals/agent-evals.md at the zip level:
hidden sheets, personal metadata, custom props, comments/threaded
comments, external links, calc chain, cached formula values, VBA, and
non-builtin defined names (allowlist per fixture).
"""

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import openpyxl

_NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def scrub_metadata(wb: openpyxl.Workbook) -> None:
    """Strip personal/source metadata from a derived seed workbook."""
    props = wb.properties
    props.creator = "eval-fixture"
    props.lastModifiedBy = "eval-fixture"
    props.title = None
    props.subject = None
    props.description = None
    props.keywords = None
    props.category = None
    props.company = None
    props.manager = None


def audit_leakage(path) -> dict:
    """Zip-level leakage audit of an emitted seed workbook."""
    out: dict = {}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        out["custom_props"] = "docProps/custom.xml" in names
        out["calc_chain"] = any(n.lower() == "xl/calcchain.xml" for n in names)
        out["vba"] = any(n.endswith("vbaProject.bin") for n in names)
        out["external_links"] = [n for n in names
                                 if n.startswith("xl/externalLinks/")]
        out["comments"] = [n for n in names
                           if n.startswith(("xl/comments",
                                            "xl/threadedComments",
                                            "xl/persons/"))]
        wb_root = ET.fromstring(z.read("xl/workbook.xml"))
        out["hidden_sheets"] = [
            sh.attrib.get("name")
            for sh in wb_root.findall(".//main:sheet", _NS)
            if sh.attrib.get("state", "visible") != "visible"]
        out["defined_names"] = [
            dn.attrib.get("name")
            for dn in wb_root.findall(".//main:definedName", _NS)]
        cached = 0
        for n in names:
            if re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                root = ET.fromstring(z.read(n))
                for c in root.findall(".//main:c", _NS):
                    f = c.find("main:f", _NS)
                    v = c.find("main:v", _NS)
                    if (f is not None and v is not None
                            and "".join(v.itertext())):
                        cached += 1
        out["cached_formula_cells"] = cached
        core = z.read("docProps/core.xml").decode("utf-8", "replace")
        out["creators"] = re.findall(
            r"<(?:dc:creator|cp:lastModifiedBy)>([^<]*)<", core)
    return out


def assert_no_leakage(path, allow_defined_names=(), allow_cached=False,
                      creator="eval-fixture") -> None:
    """Fail the build if an emitted seed violates the leakage sweep."""
    a = audit_leakage(path)
    problems = []
    if a["hidden_sheets"]:
        problems.append(f"hidden sheets: {a['hidden_sheets']}")
    if a["custom_props"]:
        problems.append("custom document properties present")
    if a["calc_chain"]:
        problems.append("calcChain.xml present")
    if a["vba"]:
        problems.append("VBA project present")
    if a["external_links"]:
        problems.append(f"external links: {a['external_links']}")
    if a["comments"]:
        problems.append(f"comments parts: {a['comments']}")
    if not allow_cached and a["cached_formula_cells"]:
        problems.append(
            f"{a['cached_formula_cells']} formula cells carry cached values")
    bad_names = [n for n in a["defined_names"]
                 if not n.startswith("_xlnm.") and n not in allow_defined_names]
    if bad_names:
        problems.append(f"defined names: {bad_names}")
    bad_creators = [c for c in a["creators"] if c != creator]
    if bad_creators:
        problems.append(f"metadata identities: {bad_creators}")
    if problems:
        raise SystemExit(
            f"leakage in {Path(path).name}: " + "; ".join(problems))
