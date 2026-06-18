#!/usr/bin/env python3
"""Audit a generated PPTX for expected labels and risky formula text."""

from __future__ import annotations

import argparse
import html
import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
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
QUESTION_RE = re.compile(r"(?:如图|下列|则|大小|方向|作用力|摩擦力|电场|磁场|速度|加速度|为\s*\(|为（|是\s*\(|是（)")
OPTION_NAME_RE = re.compile(r"(?:Options|options|Option Label)", re.I)
BODY_NAME_RE = re.compile(r"(?:Content Body|Summary Body)", re.I)
LABEL_NAME_RE = re.compile(r"(?:KEY POINT|label|Freeform|Title)", re.I)
HANDOUT_IMAGE_RE = re.compile(r"(?:^rId\d+\.(?:png|jpg|jpeg)$|(?:transparent|lecture|handout|media).*\.(?:png|jpg|jpeg)$)", re.I)

EMU_PER_INCH = 914400
DEFAULT_LAYOUT_MARGIN_IN = 0.18
PPT_NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def _emu_to_in(value: int) -> float:
    return round(value / EMU_PER_INCH, 3)


def _slide_size(zf: zipfile.ZipFile) -> tuple[int, int]:
    try:
        root = ET.fromstring(zf.read("ppt/presentation.xml"))
    except KeyError:
        return 12192000, 6858000
    size = root.find(".//p:sldSz", PPT_NS)
    if size is None:
        return 12192000, 6858000
    return int(size.get("cx", "12192000")), int(size.get("cy", "6858000"))


def _shape_bounds(elem: ET.Element) -> tuple[int, int, int, int] | None:
    xfrm = elem.find(".//a:xfrm", PPT_NS)
    if xfrm is None:
        return None
    off = xfrm.find("a:off", PPT_NS)
    ext = xfrm.find("a:ext", PPT_NS)
    if off is None or ext is None:
        return None
    return (
        int(off.get("x", "0")),
        int(off.get("y", "0")),
        int(ext.get("cx", "0")),
        int(ext.get("cy", "0")),
    )


def _bounds_area(bounds: tuple[int, int, int, int]) -> int:
    return max(0, bounds[2]) * max(0, bounds[3])


def _intersect(a: tuple[int, int, int, int], b: tuple[int, int, int, int]) -> tuple[int, int, int, int] | None:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    left = max(ax, bx)
    top = max(ay, by)
    right = min(ax + aw, bx + bw)
    bottom = min(ay + ah, by + bh)
    if right <= left or bottom <= top:
        return None
    return left, top, right - left, bottom - top


def _shape_name(elem: ET.Element) -> str:
    c_nv_pr = elem.find(".//p:cNvPr", PPT_NS)
    return c_nv_pr.get("name", "") if c_nv_pr is not None else ""


def _font_size_pt(elem: ET.Element) -> int:
    sizes: list[int] = []
    for rpr in elem.findall(".//a:rPr", PPT_NS):
        value = rpr.get("sz")
        if value:
            try:
                sizes.append(int(value) // 100)
            except ValueError:
                pass
    return max(sizes or [20])


def _line_spacing_factor(elem: ET.Element) -> float:
    for pct in elem.findall(".//a:pPr/a:lnSpc/a:spcPct", PPT_NS):
        value = pct.get("val")
        if value:
            try:
                return max(0.8, int(value) / 100000)
            except ValueError:
                pass
    return 1.2


def _text_units(text: str) -> float:
    units = 0.0
    for ch in text:
        if ch.isspace():
            units += 0.35
        elif ord(ch) < 128:
            units += 0.52
        else:
            units += 1.0
    return units


def _estimated_text_height(elem: ET.Element, bounds: tuple[int, int, int, int], text: str, paragraph_count: int) -> int:
    _x, _y, width, _height = bounds
    font_pt = _font_size_pt(elem)
    line_spacing = _line_spacing_factor(elem)
    width_pt = max(24.0, width / EMU_PER_INCH * 72)
    units_per_line = max(4.0, width_pt / max(8.0, font_pt * 0.95))
    lines = 0
    for paragraph in text.splitlines() or [text]:
        units = _text_units(paragraph)
        lines += max(1, int((units + units_per_line - 0.01) // units_per_line))
    paragraph_gap_pt = max(0, paragraph_count - 1) * 3
    height_pt = lines * font_pt * line_spacing + paragraph_gap_pt + 8
    return int(height_pt / 72 * EMU_PER_INCH)


def _text_and_breaks(elem: ET.Element) -> tuple[str, int, int]:
    paragraphs: list[str] = []
    hard_breaks = 0
    for para in elem.findall(".//a:p", PPT_NS):
        pieces: list[str] = []
        for child in para.iter():
            if child.tag == f"{{{PPT_NS['a']}}}t" and child.text:
                pieces.append(child.text)
            elif child.tag == f"{{{PPT_NS['a']}}}br":
                pieces.append("\n")
                hard_breaks += 1
        text = "".join(pieces).strip()
        if text:
            paragraphs.append(text)
    return "\n".join(paragraphs), hard_breaks, len(paragraphs)


def _warning(
    slide: int,
    kind: str,
    name: str,
    issue: str,
    bounds: tuple[int, int, int, int],
    text: str = "",
) -> dict[str, object]:
    x, y, cx, cy = bounds
    return {
        "slide": slide,
        "kind": kind,
        "name": name,
        "issue": issue,
        "x_in": _emu_to_in(x),
        "y_in": _emu_to_in(y),
        "w_in": _emu_to_in(cx),
        "h_in": _emu_to_in(cy),
        "text": text[:80],
    }


def _overlap_warning(
    slide: int,
    first: dict[str, object],
    second: dict[str, object],
    overlap: tuple[int, int, int, int],
    ratio: float,
) -> dict[str, object]:
    x, y, cx, cy = overlap
    return {
        "slide": slide,
        "kind": f"{first['kind']}-{second['kind']}",
        "name": f"{first['name']} <> {second['name']}",
        "issue": "content_overlap",
        "x_in": _emu_to_in(x),
        "y_in": _emu_to_in(y),
        "w_in": _emu_to_in(cx),
        "h_in": _emu_to_in(cy),
        "overlap_ratio": round(ratio, 3),
        "text": (str(first.get("text") or "")[:36] + " | " + str(second.get("text") or "")[:36]).strip(),
    }


def _content_near_edge(
    bounds: tuple[int, int, int, int],
    slide_w: int,
    slide_h: int,
    margin: int,
) -> bool:
    x, y, cx, cy = bounds
    return x < margin or y < margin or slide_w - (x + cx) < margin or slide_h - (y + cy) < margin


def _content_outside_slide(bounds: tuple[int, int, int, int], slide_w: int, slide_h: int) -> bool:
    x, y, cx, cy = bounds
    return x < 0 or y < 0 or x + cx > slide_w or y + cy > slide_h


def _is_template_or_label_text(shape_name: str, text: str) -> bool:
    normalized = re.sub(r"\s+", "", text)
    return bool(LABEL_NAME_RE.search(shape_name)) or bool(LABEL_RE.match(text.strip())) or len(normalized) <= 12


def _is_probable_content_image(bounds: tuple[int, int, int, int], slide_w: int, slide_h: int, pic_name: str) -> bool:
    x, y, cx, cy = bounds
    is_background = cx > slide_w * 0.92 and cy > slide_h * 0.92
    if is_background:
        return False
    return bool(HANDOUT_IMAGE_RE.search(pic_name))


def _layout_warnings(zf: zipfile.ZipFile, slide_names: list[str], margin_in: float) -> list[dict[str, object]]:
    slide_w, slide_h = _slide_size(zf)
    margin = int(margin_in * EMU_PER_INCH)
    warnings: list[dict[str, object]] = []

    for slide_index, name in enumerate(slide_names, start=1):
        root = ET.fromstring(zf.read(name))
        content_items: list[dict[str, object]] = []
        has_courseware_content = False
        has_options_text_shape = any(
            OPTION_NAME_RE.search(_shape_name(shape)) for shape in root.findall(".//p:sp", PPT_NS)
        )

        for shape in root.findall(".//p:sp", PPT_NS):
            bounds = _shape_bounds(shape)
            text, hard_breaks, paragraph_count = _text_and_breaks(shape)
            if bounds is None or not text:
                continue
            shape_name = _shape_name(shape)
            normalized = re.sub(r"\s+", "", text)
            is_question = len(normalized) >= 24 and QUESTION_RE.search(text)
            is_options = bool(OPTION_NAME_RE.search(shape_name))
            is_bullet_list = bool(re.search(r"(^|\n)\s*(?:[•●]|[0-9]+[.．、])", text))
            is_body = bool(BODY_NAME_RE.search(shape_name)) or is_bullet_list
            is_label = _is_template_or_label_text(shape_name, text)
            if is_question or is_options or is_body or QUESTION_RE.search(text):
                has_courseware_content = True

            if _content_outside_slide(bounds, slide_w, slide_h):
                warnings.append(_warning(slide_index, "text", shape_name, "text_outside_slide", bounds, text))
            elif not is_label and _content_near_edge(bounds, slide_w, slide_h, margin):
                warnings.append(_warning(slide_index, "text", shape_name, "text_near_safe_edge", bounds, text))

            if hard_breaks and len(normalized) >= 18:
                warnings.append(_warning(slide_index, "text", shape_name, "hard_line_break_in_long_text", bounds, text))
            if paragraph_count > 1 and is_question and not is_options and not is_body:
                warnings.append(_warning(slide_index, "text", shape_name, "question_split_across_paragraphs", bounds, text))

            x, _y, cx, _cy = bounds
            if is_question and not is_options and cx < slide_w * 0.55:
                warnings.append(_warning(slide_index, "text", shape_name, "question_text_box_too_narrow", bounds, text))
            if LABEL_RE.match(text.strip()) and "\n" in text:
                warnings.append(_warning(slide_index, "text", shape_name, "label_contains_line_break", bounds, text))
            is_courseware_text = is_question or is_options or is_body
            if not is_label and is_courseware_text:
                estimated_h = _estimated_text_height(shape, bounds, text, paragraph_count)
                x, y, cx, cy = bounds
                if estimated_h > cy * 1.08:
                    overflow = (x, y, cx, estimated_h)
                    item = _warning(slide_index, "text", shape_name, "text_overflow_estimated", overflow, text)
                    item["box_h_in"] = _emu_to_in(cy)
                    item["estimated_h_in"] = _emu_to_in(estimated_h)
                    warnings.append(item)
                content_items.append(
                    {
                        "kind": "text",
                        "name": shape_name,
                        "bounds": (x, y, cx, max(cy, estimated_h)),
                        "text": text,
                        "is_options": is_options,
                        "is_body": is_body,
                    }
                )

        for picture in root.findall(".//p:pic", PPT_NS):
            bounds = _shape_bounds(picture)
            if bounds is None:
                continue
            pic_name = _shape_name(picture)
            x, y, cx, cy = bounds

            is_handout_image = _is_probable_content_image(bounds, slide_w, slide_h, pic_name)
            if not is_handout_image:
                continue

            if _content_outside_slide(bounds, slide_w, slide_h):
                warnings.append(_warning(slide_index, "image", pic_name, "image_outside_slide", bounds))
            elif _content_near_edge(bounds, slide_w, slide_h, margin):
                warnings.append(_warning(slide_index, "image", pic_name, "image_near_safe_edge", bounds))

            is_wide_option_image = cx > slide_w * 0.35 and cy > slide_h * 0.12
            center_offset = abs((x + cx / 2) - slide_w / 2)
            if is_wide_option_image and not has_options_text_shape and center_offset > slide_w * 0.08:
                warnings.append(_warning(slide_index, "image", pic_name, "wide_image_not_centered", bounds))
            content_items.append({"kind": "image", "name": pic_name, "bounds": bounds, "text": ""})

        if has_courseware_content:
            for idx, first in enumerate(content_items):
                for second in content_items[idx + 1 :]:
                    overlap = _intersect(first["bounds"], second["bounds"])  # type: ignore[arg-type]
                    if overlap is None:
                        continue
                    ow, oh = overlap[2], overlap[3]
                    if ow < margin * 0.35 or oh < margin * 0.25:
                        continue
                    ratio = _bounds_area(overlap) / max(1, min(_bounds_area(first["bounds"]), _bounds_area(second["bounds"])))  # type: ignore[arg-type]
                    if ratio < 0.035:
                        continue
                    # A very wide option text box may intentionally reserve blank
                    # room for a right-side diagram. Flag only if the intersection
                    # is not confined to the far empty edge of the option box.
                    if {first["kind"], second["kind"]} == {"text", "image"}:
                        text_item = first if first["kind"] == "text" else second
                        if text_item.get("is_options") and ratio < 0.12:
                            continue
                    warnings.append(_overlap_warning(slide_index, first, second, overlap, ratio))

    return warnings


def extract_text(
    pptx: Path,
    layout_margin_in: float,
) -> tuple[str, str, list[str], list[dict[str, object]], int, int, int, int, int]:
    with zipfile.ZipFile(pptx) as zf:
        slide_names = sorted(
            [name for name in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", name)],
            key=lambda name: int(re.search(r"slide(\d+)\.xml$", name).group(1)),
        )
        slide_xml = [zf.read(name).decode("utf-8", errors="ignore") for name in slide_names]
        xml = "\n".join(slide_xml)
        media_count = len([name for name in zf.namelist() if name.startswith("ppt/media/")])
        layout_warnings = _layout_warnings(zf, slide_names, layout_margin_in)
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
        layout_warnings,
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
    parser.add_argument(
        "--layout-margin-in",
        type=float,
        default=DEFAULT_LAYOUT_MARGIN_IN,
        help="Safe-edge margin in inches for text and content images.",
    )
    parser.add_argument(
        "--fail-layout",
        action="store_true",
        help="Fail when layout warnings are found.",
    )
    parser.add_argument("--json", action="store_true", help="Print JSON only.")
    args = parser.parse_args()

    (
        text,
        raw_xml,
        plain_formula_text,
        layout_warnings,
        slide_count,
        media_count,
        image_placements,
        formula_object_count,
        radical_object_count,
    ) = extract_text(args.pptx, args.layout_margin_in)
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
        "layout_margin_in": args.layout_margin_in,
        "layout_warnings": layout_warnings,
        "media_count": media_count,
        "image_placements": image_placements,
        "passed": not missing
        and not found_bad
        and not raw_sqrt_without_radical
        and (args.allow_plain_formulas or not plain_formula_text)
        and (not args.fail_layout or not layout_warnings),
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
        if layout_warnings:
            print(f"Layout warnings: {len(layout_warnings)}")
            for item in layout_warnings[:40]:
                text_preview = item.get("text") or ""
                suffix = f" | {text_preview}" if text_preview else ""
                print(
                    "  - "
                    f"slide {item['slide']} {item['issue']} "
                    f"({item['kind']}, {item['name']}, "
                    f"x={item['x_in']}in y={item['y_in']}in "
                    f"w={item['w_in']}in h={item['h_in']}in)"
                    f"{suffix}"
                )
            if len(layout_warnings) > 40:
                print(f"  ... {len(layout_warnings) - 40} more")
        print("PASS" if report["passed"] else "FAIL")

    return 0 if report["passed"] else 1


if __name__ == "__main__":
    sys.exit(main())
