#!/usr/bin/env python3
"""Audit a generated PPTX for expected labels and risky formula text."""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
import zipfile
from pathlib import Path


DEFAULT_BAD_TOKENS = [
    "L1-",
    "sqrt",
    "表达式A",
    "表达式B",
    "表达式C",
    "表达式D",
    "q1 q2",
    "r^2",
    "...",
]


def extract_text(pptx: Path) -> tuple[str, int, int, int]:
    with zipfile.ZipFile(pptx) as zf:
        slide_names = sorted(
            [name for name in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", name)],
            key=lambda name: int(re.search(r"slide(\d+)\.xml$", name).group(1)),
        )
        xml = "\n".join(zf.read(name).decode("utf-8", errors="ignore") for name in slide_names)
        media_count = len([name for name in zf.namelist() if name.startswith("ppt/media/")])
    texts = [html.unescape(match.group(1)) for match in re.finditer(r"<a:t>(.*?)</a:t>", xml)]
    image_placements = xml.count("<a:blip ")
    return "\n".join(texts), len(slide_names), media_count, image_placements


def load_labels(path: Path | None) -> list[str]:
    if not path:
        return []
    return [
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip() and not line.strip().startswith("#")
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description="Audit PPTX labels and risky formula text.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--expect-labels", type=Path, help="Text file with one expected label per line.")
    parser.add_argument("--bad-token", action="append", default=[], help="Additional bad token to scan for.")
    parser.add_argument("--json", action="store_true", help="Print JSON only.")
    args = parser.parse_args()

    text, slide_count, media_count, image_placements = extract_text(args.pptx)
    expected = load_labels(args.expect_labels)
    bad_tokens = DEFAULT_BAD_TOKENS + args.bad_token
    missing = [label for label in expected if label not in text]
    found_bad = [token for token in bad_tokens if token in text]

    report = {
        "pptx": str(args.pptx),
        "slide_count": slide_count,
        "expected_labels": len(expected),
        "labels_found": len(expected) - len(missing),
        "missing_labels": missing,
        "bad_tokens": found_bad,
        "media_count": media_count,
        "image_placements": image_placements,
        "passed": not missing and not found_bad,
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"PPTX: {report['pptx']}")
        print(f"Slides: {slide_count}")
        print(f"Labels: {report['labels_found']}/{report['expected_labels']}")
        print(f"Media: {media_count}, image placements: {image_placements}")
        if missing:
            print("Missing labels:")
            for label in missing:
                print(f"  - {label}")
        if found_bad:
            print("Risky tokens:")
            for token in found_bad:
                print(f"  - {token}")
        print("PASS" if report["passed"] else "FAIL")

    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
