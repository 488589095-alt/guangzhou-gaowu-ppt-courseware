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
    "√",
    "表达式A",
    "表达式B",
    "表达式C",
    "表达式D",
    "q1 q2",
    "r^2",
    "...",
]

FORMULA_TEXT_RE = re.compile(
    r"(?:F\s*=|E\s*=|kq|kQ|mg(?:tan|cos)|πkσ|q₁q₂|N·m²/C²|/[0-9]?[A-Za-zLRdrql₀θ²³₁₂₃₄₅₆₇₈₉₀]+)"
)
LABEL_RE = re.compile(r"^(?:P\d+[-－](?:例|练|巩固)\d+(?:-\d+)?|L\d+[-－](?:例|练)\d+|[A-D]\.?)$")


def extract_text(pptx: Path) -> tuple[str, str, list[str], int, int, int, int, int]:
    with zipfile.ZipFile(pptx) as zf:
        slide_names = sorted(
            [name for name in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", name)],
            key=lambda name: int(re.search(r"slide(\d+)\.xml$", name).group(1)),
        )
        slide_xml = [zf.read(name).decode("utf-8", errors="ignore") for name in slide_names]
        xml = "\n".join(slide_xml)
        media_count = len([name for name in zf.namelist() if name.startswith("ppt/media/")])
    texts = [html.unescape(match.group(1)) for match in re.finditer(r"<a:t>(.*?)</a:t>", xml)]
    math_texts = [html.unescape(match.group(1)) for match in re.finditer(r"<m:t>(.*?)</m:t>", xml)]
    all_text = "\n".join(texts + math_texts)
    image_placements = xml.count("<a:blip ")
    formula_object_count = xml.count("<m:oMath")
    radical_object_count = xml.count("<m:rad")
    plain_formula_text = [
        text.strip()
        for text in texts
        if text.strip()
        and not LABEL_RE.match(text.strip())
        and FORMULA_TEXT_RE.search(text.strip())
    ]
    return (
        all_text,
        xml,
        plain_formula_text,
        len(slide_names),
        media_count,
        image_placements,
        formula_object_count,
        radical_object_count,
    )


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
    parser.add_argument(
        "--allow-plain-formulas",
        action="store_true",
        help="Do not fail when formula-like strings remain in normal text runs.",
    )
    parser.add_argument("--json", action="store_true", help="Print JSON only.")
    args = parser.parse_args()

    (
        text,
        raw_xml,
        plain_formula_text,
        slide_count,
        media_count,
        image_placements,
        formula_object_count,
        radical_object_count,
    ) = extract_text(args.pptx)
    expected = load_labels(args.expect_labels)
    bad_tokens = DEFAULT_BAD_TOKENS + args.bad_token
    missing = [label for label in expected if label not in text]
    found_bad = [token for token in bad_tokens if token in text]
    raw_sqrt_count = raw_xml.count("√")
    raw_sqrt_without_radical = raw_sqrt_count > 0

    report = {
        "pptx": str(args.pptx),
        "slide_count": slide_count,
        "expected_labels": len(expected),
        "labels_found": len(expected) - len(missing),
        "missing_labels": missing,
        "bad_tokens": found_bad,
        "raw_sqrt_count": raw_sqrt_count,
        "formula_objects": formula_object_count,
        "radical_objects": radical_object_count,
        "plain_formula_text": plain_formula_text,
        "media_count": media_count,
        "image_placements": image_placements,
        "passed": not missing
        and not found_bad
        and not raw_sqrt_without_radical
        and (args.allow_plain_formulas or not plain_formula_text),
    }

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print(f"PPTX: {report['pptx']}")
        print(f"Slides: {slide_count}")
        print(f"Labels: {report['labels_found']}/{report['expected_labels']}")
        print(f"Media: {media_count}, image placements: {image_placements}")
        print(f"Formula objects: {formula_object_count}, radical objects: {radical_object_count}")
        if raw_sqrt_count:
            print(f"Raw sqrt chars in XML: {raw_sqrt_count}")
        if missing:
            print("Missing labels:")
            for label in missing:
                print(f"  - {label}")
        if found_bad:
            print("Risky tokens:")
            for token in found_bad:
                print(f"  - {token}")
        if plain_formula_text and not args.allow_plain_formulas:
            print("Formula-like text still in normal text runs:")
            for item in plain_formula_text[:30]:
                print(f"  - {item}")
            if len(plain_formula_text) > 30:
                print(f"  ... {len(plain_formula_text) - 30} more")
        print("PASS" if report["passed"] else "FAIL")

    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
