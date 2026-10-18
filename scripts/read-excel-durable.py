#!/usr/bin/env python3
"""Independently inspect a scratch Durable database persisted by Excel's WebKit.

Reads SQLite directly, not the pane or bridge. Only databases with the Durable
prefix are accepted. The default output lists database names and object stores;
--database selects one exact name and decodes its records. Output may contain
workbook contents and conversation text: use scratch workbooks, not customer data.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sqlite3
import sys

ROOT = Path(__file__).resolve().parent.parent
DECODER = ROOT / ".agents/skills/excel-background-verification/scripts/read-taskpane-setting.py"
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/Default"
PREFIX = "pi-for-excel-durable:"


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--database", help="exact Durable database name from the bridge or listing")
    args = parser.parse_args()
    if args.database and not args.database.startswith(PREFIX):
        parser.error(f"only {PREFIX} databases can be inspected")
    spec = importlib.util.spec_from_file_location("webkit_decoder", DECODER)
    if spec is None or spec.loader is None:
        sys.exit("cannot load the WebKit record decoder")
    sys.dont_write_bytecode = True
    decoder_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(decoder_module)
    output = []
    # pathlib.glob suppresses PermissionError and would falsely report no data
    # when macOS privacy permissions deny the Excel container. Probe access
    # explicitly before enumerating the known WebKit database layout.
    try:
        next(CONTAINER.iterdir(), None)
    except FileNotFoundError:
        print("[]")
        return
    except PermissionError:
        sys.exit("Cannot inspect Excel WebKit storage: filesystem permission was denied. No database absence can be inferred.")
    for path in sorted(CONTAINER.glob("*/*/IndexedDB/*/IndexedDB.sqlite3")):
        with sqlite3.connect(f"{path.as_uri()}?mode=ro", uri=True) as source:
            info = dict(source.execute("select key, value from IDBDatabaseInfo"))
            raw_name = info.get("DatabaseName", "")
            name = raw_name.decode("utf-16-le") if isinstance(raw_name, bytes) else raw_name
            if not name.startswith(PREFIX) or (args.database and name != args.database):
                continue
            # SQLite backup observes committed WAL content without copying files
            # while Excel is writing or modifying the live database. Never copy
            # the ordinary pane database, which also contains provider secrets.
            with sqlite3.connect(":memory:") as snapshot:
                source.backup(snapshot)
                stores = snapshot.execute("select id, name from ObjectStoreInfo order by id").fetchall()
                item = {"database": name, "stores": {}}
                for store_id, store_name in stores:
                    rows = snapshot.execute("select cast(key as blob), value from Records where objectStoreID = ?", (store_id,)).fetchall()
                    if not args.database:
                        item["stores"][store_name] = {"records": len(rows)}
                        continue
                    records = []
                    for key, value in rows:
                        decoder = decoder_module.Decoder(bytes(value))
                        version = decoder.u32()
                        if version != decoder_module.SERIALIZATION_VERSION:
                            sys.exit(f"unsupported WebKit serialized value version: {version}")
                        raw_key = key.encode("latin-1") if isinstance(key, str) else bytes(key)
                        label = decoder_module.decode_key(key) if len(raw_key) >= 2 and raw_key[1] == 0x60 else raw_key.hex()
                        records.append({"key": label, "value": decoder.value()})
                    item["stores"][store_name] = records
                output.append(item)
    if args.database and not output:
        sys.exit("requested Durable database was not found")
    print(json.dumps(output, indent=2))


if __name__ == "__main__":
    main()
