#!/usr/bin/env python3
# From the product-description skill: https://gist.github.com/steveruizok/83ae5c53f2784ebf8f5fe0a3fb94480f
"""Check relative links and heading anchors across a product description repo.

Usage: python3 check-links.py [repo-root]   (default: current directory)

Walks every .md file, extracts every `[text](target)` and `<target>`-free
relative link, and reports targets whose file does not exist or whose
`#anchor` does not match a heading slug in the target file. Skips http(s),
mailto, and fenced code blocks. Exits 1 if anything is broken.

Slugs follow GitHub's rule: lowercase, strip everything except letters,
digits, spaces, and hyphens, spaces to hyphens, duplicates get -1, -2, ...
"""
import os
import re
import sys
import unicodedata

ROOT = os.path.abspath(sys.argv[1] if len(sys.argv) > 1 else ".")
LINK = re.compile(r"(?<!\!)\[[^\]]*\]\(([^)\s]+)(?:\s+\"[^\"]*\")?\)")
HEADING = re.compile(r"^(#{1,6})\s+(.*?)\s*#*\s*$")
FENCE = re.compile(r"^\s*(```|~~~)")


def slugify(text):
    text = re.sub(r"`([^`]*)`", r"\1", text)            # drop inline code ticks
    text = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", text)  # drop link targets
    text = re.sub(r"[*_]", "", text)                        # drop emphasis
    text = unicodedata.normalize("NFKD", text).lower()
    text = "".join(ch for ch in text if ch.isalnum() or ch in " -")
    return text.replace(" ", "-")


def strip_fences(lines):
    out, in_fence = [], False
    for line in lines:
        if FENCE.match(line):
            in_fence = not in_fence
            out.append("")
            continue
        out.append("" if in_fence else line)
    return out


def headings_of(path):
    with open(path, encoding="utf-8") as f:
        lines = strip_fences(f.read().splitlines())
    seen, slugs = {}, set()
    for line in lines:
        m = HEADING.match(line)
        if not m:
            continue
        base = slugify(m.group(2))
        n = seen.get(base, 0)
        seen[base] = n + 1
        slugs.add(base if n == 0 else f"{base}-{n}")
    return slugs


def md_files():
    for dirpath, dirnames, filenames in os.walk(ROOT):
        dirnames[:] = [d for d in dirnames if not d.startswith(".") and d != "node_modules"]
        for name in filenames:
            if name.endswith(".md"):
                yield os.path.join(dirpath, name)


def main():
    files = sorted(md_files())
    anchors = {}
    problems = []
    checked = 0
    for path in files:
        with open(path, encoding="utf-8") as f:
            lines = strip_fences(f.read().splitlines())
        for lineno, line in enumerate(lines, 1):
            for target in LINK.findall(line):
                if re.match(r"^[a-z][a-z0-9+.-]*:", target):  # http:, mailto:, etc.
                    continue
                checked += 1
                file_part, _, anchor = target.partition("#")
                if file_part:
                    dest = os.path.normpath(os.path.join(os.path.dirname(path), file_part))
                else:
                    dest = path
                rel = os.path.relpath(path, ROOT)
                if not os.path.exists(dest):
                    problems.append(f"{rel}:{lineno}: missing file: {target}")
                    continue
                if anchor and dest.endswith(".md"):
                    if dest not in anchors:
                        anchors[dest] = headings_of(dest)
                    if anchor.lower() not in anchors[dest]:
                        problems.append(f"{rel}:{lineno}: missing anchor: {target}")
    for p in problems:
        print(p)
    print(f"{len(files)} files, {checked} relative links, {len(problems)} broken")
    sys.exit(1 if problems else 0)


if __name__ == "__main__":
    main()
