#!/usr/bin/env python3
"""Create (or tag) a scratch workbook that opens the sideloaded Pi taskpane by itself.

Excel for Mac does not put a sideloaded add-in on the ribbon until the add-in has
been activated once in that Excel profile, and the Home-tab "Add-ins" flyout does
not render for a background app. A workbook can carry an auto-open task pane part
(`xl/webextensions`); when it references the sideloaded manifest with
`store="developer" storeType="Registry"`, Excel opens the taskpane on document
open without needing the foreground. After that first activation the "Open Pi"
ribbon button exists and responds to a semantic AXPress.

Usage:
  scratch-workbook.py create <path.xlsx> [--addin-id ID]
  scratch-workbook.py tag <existing.xlsx> [--addin-id ID]

`create` writes a minimal one-sheet workbook plus the auto-open part. `tag` adds
the part to an existing workbook in place. Save the workbook inside the Excel
container (~/Library/Containers/com.microsoft.Excel/Data/Documents/) so the
sandbox does not raise a "Grant File Access" sheet.

The add-in ID defaults to the <Id> in the repository manifest.xml.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
import uuid
import zipfile
from pathlib import Path

CONTENT_TYPES_OVERRIDES = (
    '<Override PartName="/xl/webextensions/taskpanes.xml" '
    'ContentType="application/vnd.ms-office.webextensiontaskpanes+xml"/>'
    '<Override PartName="/xl/webextensions/webextension1.xml" '
    'ContentType="application/vnd.ms-office.webextension+xml"/>'
)
ROOT_REL = (
    '<Relationship Id="rIdPiTaskpanes" '
    'Type="http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes" '
    'Target="xl/webextensions/taskpanes.xml"/>'
)
TASKPANES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<wetp:taskpanes xmlns:wetp="http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11">'
    '<wetp:taskpane dockstate="right" visibility="1" width="420" row="1">'
    '<wetp:webextensionref xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>'
    "</wetp:taskpane></wetp:taskpanes>"
)
TASKPANES_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/webextension" '
    'Target="webextension1.xml"/></Relationships>'
)

MINIMAL_WORKBOOK = {
    "[Content_Types].xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    ),
    "_rels/.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/></Relationships>'
    ),
    "xl/workbook.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    "xl/_rels/workbook.xml.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        'Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    "xl/worksheets/sheet1.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        "<sheetData/></worksheet>"
    ),
}


def default_addin_id() -> str:
    manifest = Path(__file__).resolve().parents[4] / "manifest.xml"
    match = re.search(r"<Id>([^<]+)</Id>", manifest.read_text())
    if not match:
        sys.exit(f"could not find <Id> in {manifest}")
    return match.group(1)


def webextension_xml(addin_id: str) -> str:
    # Excel keys document-level add-in instances by this GUID within a running
    # process; a fresh one per workbook keeps a second scratch workbook from
    # being treated as the already-open instance.
    instance_id = "{" + str(uuid.uuid4()).upper() + "}"
    reference = f'<we:reference id="{addin_id}" version="1.0.0.0" store="developer" storeType="Registry"/>'
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<we:webextension xmlns:we="http://schemas.microsoft.com/office/webextensions/webextension/2010/11" '
        f'id="{instance_id}">{reference}'
        "<we:alternateReferences/><we:properties/><we:bindings/>"
        '<we:snapshot xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>'
        "</we:webextension>"
    )


def with_taskpane_parts(parts: dict[str, bytes], addin_id: str) -> dict[str, bytes]:
    out = {name: data for name, data in parts.items() if not name.startswith("xl/webextensions/")}
    content_types = out["[Content_Types].xml"].decode()
    if "webextensiontaskpanes" not in content_types:
        out["[Content_Types].xml"] = content_types.replace("</Types>", CONTENT_TYPES_OVERRIDES + "</Types>").encode()
    rels = out["_rels/.rels"].decode()
    if "webextensiontaskpanes" not in rels:
        out["_rels/.rels"] = rels.replace("</Relationships>", ROOT_REL + "</Relationships>").encode()
    out["xl/webextensions/taskpanes.xml"] = TASKPANES_XML.encode()
    out["xl/webextensions/_rels/taskpanes.xml.rels"] = TASKPANES_RELS.encode()
    out["xl/webextensions/webextension1.xml"] = webextension_xml(addin_id).encode()
    return out


def write_workbook(path: Path, parts: dict[str, bytes]) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts.items():
            zout.writestr(name, data)
    os.replace(tmp, path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("mode", choices=("create", "tag"))
    parser.add_argument("path", type=Path)
    parser.add_argument("--addin-id", default=None)
    args = parser.parse_args()
    addin_id = args.addin_id or default_addin_id()

    if args.mode == "create":
        parts = {name: data.encode() for name, data in MINIMAL_WORKBOOK.items()}
    else:
        with zipfile.ZipFile(args.path) as zin:
            parts = {item.filename: zin.read(item.filename) for item in zin.infolist()}

    write_workbook(args.path, with_taskpane_parts(parts, addin_id))
    print(f"{args.mode}: {args.path} opens add-in {addin_id} on document open")


if __name__ == "__main__":
    main()
