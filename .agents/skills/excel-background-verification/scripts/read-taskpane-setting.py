#!/usr/bin/env python3
"""Read a persisted taskpane setting straight from Excel's WebKit IndexedDB.

The taskpane stores settings in the `settings` object store of its `pi-for-excel`
IndexedDB database. Excel for Mac keeps that database at
~/Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/…/IndexedDB/…/IndexedDB.sqlite3
(plus a WAL that holds the latest writes). This script copies the database files
and decodes one record outside the app, so a verification run can assert what
was persisted independently of the pane's own code.

Usage:
  read-taskpane-setting.py <key> [--list]

Prints JSON. Records are WebKit SerializedScriptValue v15; the decoder covers
the subset the settings store writes (objects, arrays, strings, numbers,
booleans, null/undefined, dates) and the string constant pool.
"""

from __future__ import annotations

import argparse
from contextlib import closing
import json
import os
import shutil
import sqlite3
import struct
import sys
import tempfile
from pathlib import Path

CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/Library/WebKit/WebsiteData/Default"
SERIALIZATION_VERSION = 15
STRING_POOL_MARKER = 0xFFFFFFFE
TERMINATOR = 0xFFFFFFFF
STRING_IS_8BIT = 0x80000000


class Decoder:
    def __init__(self, data: bytes) -> None:
        self.data = data
        self.pos = 0
        self.pool: list[str] = []

    def u8(self) -> int:
        value = self.data[self.pos]
        self.pos += 1
        return value

    def u32(self) -> int:
        value = struct.unpack_from("<I", self.data, self.pos)[0]
        self.pos += 4
        return value

    def string(self, header: int) -> str:
        if header == STRING_POOL_MARKER:
            size = len(self.pool)
            if size <= 0xFF:
                index = self.u8()
            elif size <= 0xFFFF:
                index = struct.unpack_from("<H", self.data, self.pos)[0]
                self.pos += 2
            else:
                index = self.u32()
            return self.pool[index]
        length = header & ~STRING_IS_8BIT
        if header & STRING_IS_8BIT:
            text = self.data[self.pos : self.pos + length].decode("latin-1")
            self.pos += length
        else:
            text = self.data[self.pos : self.pos + length * 2].decode("utf-16-le")
            self.pos += length * 2
        self.pool.append(text)
        return text

    def properties(self, into: dict[str, object]) -> None:
        while True:
            header = self.u32()
            if header == TERMINATOR:
                return
            name = self.string(header)
            into[name] = self.value()

    def value(self) -> object:
        tag = self.u8()
        if tag == 0x01:  # ArrayTag
            out: list[object] = [None] * self.u32()
            while True:
                index = self.u32()
                if index == TERMINATOR:
                    break
                out[index] = self.value()
            self.properties({})
            return out
        if tag == 0x02:  # ObjectTag
            obj: dict[str, object] = {}
            self.properties(obj)
            return obj
        if tag in (0x03, 0x04):  # Undefined, Null
            return None
        if tag == 0x05:  # IntTag
            value = struct.unpack_from("<i", self.data, self.pos)[0]
            self.pos += 4
            return value
        if tag == 0x06:
            return 0
        if tag == 0x07:
            return 1
        if tag == 0x08:
            return False
        if tag == 0x09:
            return True
        if tag in (0x0A, 0x0B):  # Double, Date
            value = struct.unpack_from("<d", self.data, self.pos)[0]
            self.pos += 8
            return value
        if tag == 0x10:  # StringTag
            return self.string(self.u32())
        if tag == 0x11:  # EmptyStringTag
            return ""
        raise ValueError(f"unhandled serialization tag 0x{tag:02x} at offset {self.pos - 1}")


def decode_key(raw: str | bytes) -> str:
    data = raw.encode("latin-1") if isinstance(raw, str) else bytes(raw)
    # WebKit IDB key: 0x00, type byte 0x60 (string), u32 length, UTF-16LE characters.
    if len(data) >= 6 and data[1] == 0x60:
        length = struct.unpack_from("<I", data, 2)[0]
        return data[6 : 6 + length * 2].decode("utf-16-le")
    return data.decode("latin-1")


def find_database() -> Path:
    def fail_walk(error: OSError) -> None:
        raise error

    candidates: list[Path] = []
    # Unlike pathlib.glob, walk propagates denied access, including nested origins.
    # Keep traversal bounded to */*/IndexedDB/*/IndexedDB.sqlite3.
    for root, directories, files in os.walk(CONTAINER, onerror=fail_walk):
        directory = Path(root)
        depth = len(directory.relative_to(CONTAINER).parts)
        if depth == 2:
            directories[:] = [name for name in directories if name == "IndexedDB"]
        if depth == 4:
            directories.clear()
        if depth != 4 or "IndexedDB.sqlite3" not in files:
            continue
        path = directory / "IndexedDB.sqlite3"
        with closing(sqlite3.connect(f"{path.as_uri()}?mode=ro", uri=True)) as connection:
            row = connection.execute("select value from IDBDatabaseInfo where key = 'DatabaseName'").fetchone()
            raw_name = row[0] if row else None
            name = raw_name.decode("utf-16-le") if isinstance(raw_name, bytes) else raw_name
            if name != "pi-for-excel":
                continue
            store = connection.execute("select id from ObjectStoreInfo where name = 'settings'").fetchone()
            if store is not None:
                candidates.append(path)
    if not candidates:
        sys.exit("no `pi-for-excel` IndexedDB database with a `settings` object store; has the taskpane run in Excel?")
    # Retain the helper's newest-file policy across origins, but only after exact
    # app identity + store qualification. A newer Durable database is irrelevant.
    return max(candidates, key=lambda path: (path.stat().st_mtime, str(path)))


def copy_database() -> Path:
    try:
        source = find_database()
        scratch = Path(tempfile.mkdtemp(prefix="pi-idb-"))
        try:
            for suffix in ("", "-wal", "-shm"):
                candidate = source.with_name(source.name + suffix)
                if candidate.exists():
                    shutil.copy2(candidate, scratch / candidate.name)
        except BaseException:
            shutil.rmtree(scratch, ignore_errors=True)
            raise
        return scratch / source.name
    except PermissionError:
        sys.exit("Cannot inspect Excel WebKit storage: filesystem permission was denied. No database absence can be inferred.")
    except FileNotFoundError:
        sys.exit("Excel WebKit storage was not found or changed during inspection; has the taskpane run in Excel?")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("key", nargs="?")
    parser.add_argument("--list", action="store_true", help="list setting keys instead of decoding one")
    args = parser.parse_args()
    if not args.key and not args.list:
        parser.error("a key is required unless --list is given")

    database = copy_database()
    try:
        with closing(sqlite3.connect(database)) as connection:
            store = connection.execute("select id from ObjectStoreInfo where name = 'settings'").fetchone()
            if store is None:
                sys.exit("no `settings` object store in the taskpane database")
            rows = connection.execute("select key, value from Records where objectStoreID = ?", (store[0],)).fetchall()
    finally:
        shutil.rmtree(database.parent, ignore_errors=True)

    if args.list:
        print(json.dumps(sorted(decode_key(key) for key, _ in rows), indent=1))
        return

    for key, value in rows:
        if decode_key(key) != args.key:
            continue
        decoder = Decoder(bytes(value))
        version = decoder.u32()
        if version != SERIALIZATION_VERSION:
            sys.exit(f"unexpected SerializedScriptValue version {version}")
        print(json.dumps(decoder.value(), indent=1))
        return
    sys.exit(f"no setting named {args.key!r}")


if __name__ == "__main__":
    main()
