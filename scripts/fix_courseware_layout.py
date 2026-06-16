#!/usr/bin/env python3
"""Root-cause layout fixer for Gaowu courseware PPTX files.

The fixer is intentionally conservative: it only touches generated content
shapes (problem stems, options, content bodies and extracted handout images),
while leaving template background/decorative assets in place.
"""

from __future__ import annotations

import argparse
import copy
import math
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree


P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
A14 = "http://schemas.microsoft.com/office/drawing/2010/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PR = "http://schemas.openxmlformats.org/package/2006/relationships"
NS = {"p": P, "a": A, "a14": A14, "m": M, "r": R, "pr": PR}

EMU_PER_IN = 914400
OPTION_LINE_SPACING = 150000
OPTION_SPACE_AFTER = 1000

QUESTION_NAME_RE = re.compile(r"(?:Problem Stem|stem)", re.I)
OPTION_NAME_RE = re.compile(r"(?:Options|options)", re.I)
CONTENT_NAME_RE = re.compile(r"(?:Content Body|Summary Body|Text)", re.I)
LABEL_NAME_RE = re.compile(r"(?:KEY POINT|label|Option Label|Freeform)", re.I)
OPTION_PREFIX_RE = re.compile(r"^([A-D][．.]\s*)(.+)$")
FORMULA_ONLY_RE = re.compile(r"^[\sA-Za-z0-9₀₁₂₃₄₅₆₇₈₉ₐᵦαβθπσμΔδFNGQklrabvxyzXYZRLO＋+\-*/=<>≤≥（）()\[\]{}√²³^·.,，]+$")
FORMULA_RISK_RE = re.compile(r"(√|/|²|³|₀|₁|₂|₃|₄|₅|₆|₇|₈|₉|[A-Za-z]\s*=)")


def q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def emu(inches: float) -> int:
    return int(round(inches * EMU_PER_IN))


def to_in(value: int) -> float:
    return value / EMU_PER_IN


def parse(data: bytes):
    return etree.fromstring(data)


def xml_bytes(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def slide_size(entries: dict[str, bytes]) -> tuple[int, int]:
    root = parse(entries["ppt/presentation.xml"])
    node = root.find(".//p:sldSz", NS)
    if node is None:
        return 12192000, 6858000
    return int(node.get("cx", "12192000")), int(node.get("cy", "6858000"))


def shape_name(el) -> str:
    node = el.find(".//p:cNvPr", NS)
    return node.get("name", "") if node is not None else ""


def shape_bounds(el) -> tuple[int, int, int, int] | None:
    xfrm = el.find(".//a:xfrm", NS)
    if xfrm is None:
        return None
    off = xfrm.find("a:off", NS)
    ext = xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    return (
        int(off.get("x", "0")),
        int(off.get("y", "0")),
        int(ext.get("cx", "0")),
        int(ext.get("cy", "0")),
    )


def set_bounds(el, x: int, y: int, w: int, h: int) -> None:
    xfrm = el.find(".//a:xfrm", NS)
    if xfrm is None:
        sppr = el.find("p:spPr", NS)
        if sppr is None:
            sppr = etree.SubElement(el, q(P, "spPr"))
        xfrm = etree.SubElement(sppr, q(A, "xfrm"))
    off = xfrm.find("a:off", NS)
    if off is None:
        off = etree.SubElement(xfrm, q(A, "off"))
    ext = xfrm.find("a:ext", NS)
    if ext is None:
        ext = etree.SubElement(xfrm, q(A, "ext"))
    off.set("x", str(int(x)))
    off.set("y", str(int(y)))
    ext.set("cx", str(int(w)))
    ext.set("cy", str(int(h)))


def text_paragraphs(sp) -> list[str]:
    out: list[str] = []
    for p_el in sp.findall(".//a:p", NS):
        parts: list[str] = []
        for child in p_el.iter():
            if child.tag == q(A, "t") and child.text:
                parts.append(child.text)
            elif child.tag == q(A, "br"):
                parts.append("\n")
        text = "".join(parts).strip()
        if text:
            out.append(text)
    return out


def text_content(sp, sep: str = "\n") -> str:
    return sep.join(text_paragraphs(sp)).strip()


def first_run_style(sp) -> dict[str, object]:
    rpr = sp.find(".//a:rPr", NS)
    style: dict[str, object] = {
        "size": 22,
        "color": "555A64",
        "bold": True,
        "font": "微软雅黑",
    }
    if rpr is not None:
        if rpr.get("sz"):
            try:
                style["size"] = max(8, int(rpr.get("sz")) // 100)
            except ValueError:
                pass
        style["bold"] = rpr.get("b") == "1"
        color = rpr.find(".//a:srgbClr", NS)
        if color is not None and color.get("val"):
            style["color"] = color.get("val")
        font = rpr.find("a:ea", NS)
        if font is None:
            font = rpr.find("a:latin", NS)
        if font is not None and font.get("typeface"):
            style["font"] = font.get("typeface")
    return style


def solid_fill(color: str):
    fill = etree.Element(q(A, "solidFill"))
    etree.SubElement(fill, q(A, "srgbClr")).set("val", color)
    return fill


def make_run(text: str, style: dict[str, object]):
    r = etree.Element(q(A, "r"))
    rpr = etree.SubElement(r, q(A, "rPr"))
    rpr.set("lang", "zh-CN")
    rpr.set("altLang", "en-US")
    rpr.set("sz", str(int(style.get("size", 22)) * 100))
    if style.get("bold", True):
        rpr.set("b", "1")
    rpr.append(solid_fill(str(style.get("color", "555A64"))))
    font = str(style.get("font", "微软雅黑"))
    etree.SubElement(rpr, q(A, "latin")).set("typeface", font)
    etree.SubElement(rpr, q(A, "ea")).set("typeface", font)
    etree.SubElement(rpr, q(A, "cs")).set("typeface", font)
    t = etree.SubElement(r, q(A, "t"))
    t.text = text
    return r


def make_ppr(line_spacing: int | None = None, align: str | None = None):
    ppr = etree.Element(q(A, "pPr"))
    if align:
        ppr.set("algn", align)
    if line_spacing:
        ln = etree.SubElement(ppr, q(A, "lnSpc"))
        etree.SubElement(ln, q(A, "spcPct")).set("val", str(line_spacing))
    return ppr


def add_spacing(ppr, before: int | None = None, after: int | None = None) -> None:
    for tag in ("spcBef", "spcAft"):
        old = ppr.find(f"a:{tag}", NS)
        if old is not None:
            ppr.remove(old)
    if before is not None:
        bef = etree.SubElement(ppr, q(A, "spcBef"))
        etree.SubElement(bef, q(A, "spcPts")).set("val", str(before))
    if after is not None:
        aft = etree.SubElement(ppr, q(A, "spcAft"))
        etree.SubElement(aft, q(A, "spcPts")).set("val", str(after))


def replace_paragraphs(sp, paragraphs) -> None:
    tx = sp.find("p:txBody", NS)
    if tx is None:
        tx = etree.SubElement(sp, q(P, "txBody"))
        body = etree.SubElement(tx, q(A, "bodyPr"))
        body.set("wrap", "square")
        body.set("anchor", "t")
        etree.SubElement(tx, q(A, "lstStyle"))
    body = tx.find("a:bodyPr", NS)
    if body is None:
        body = etree.Element(q(A, "bodyPr"))
        tx.insert(0, body)
    body.set("wrap", "square")
    body.set("anchor", "t")
    for child in list(tx):
        if child.tag == q(A, "p"):
            tx.remove(child)
    for p_el in paragraphs:
        tx.append(p_el)


def text_units(text: str) -> float:
    total = 0.0
    for ch in text:
        if ch.isspace():
            total += 0.35
        elif ord(ch) < 128:
            total += 0.52
        else:
            total += 1.0
    return total


def estimated_text_height(text: str, width: int, pt: int, line_spacing: float = 1.2) -> int:
    width_in = max(1.0, to_in(width))
    chars_per_line = max(8.0, width_in * (72 / max(pt, 12)) * 1.95)
    lines = max(1, math.ceil(text_units(text) / chars_per_line))
    return emu(lines * (pt / 72) * line_spacing + 0.14)


def normalize_sentence(text: str) -> str:
    text = re.sub(r"[\r\n]+", "", text)
    text = re.sub(r"\s{2,}", " ", text)
    text = text.replace("（ )", "（    ）").replace("( )", "（    ）")
    text = re.sub(r"（\s+）", "（    ）", text)
    text = text.replace("（\u3000\u3000 / ）", "（\u3000\u3000）")
    return text.strip()


def normalize_options(paragraphs: list[str]) -> list[str]:
    text = "\n".join(paragraphs)
    text = text.replace("\r", "\n")
    text = re.sub(r"\n+(?=[A-D][．.])", "\n", text)
    text = re.sub(r"\n+", "", text) if not re.search(r"(?:^|\n)[A-D][．.]", text) else text
    parts = re.split(r"(?=(?:^|\n)[A-D][．.])", text)
    options = [normalize_sentence(part) for part in parts if normalize_sentence(part)]
    if len(options) >= 2:
        return options
    return [normalize_sentence(p) for p in paragraphs if normalize_sentence(p)]


def normalize_bullets(paragraphs: list[str]) -> list[str]:
    out: list[str] = []
    for raw in paragraphs:
        p = raw.strip()
        if not p:
            continue
        is_new = bool(re.match(r"^(?:[•●]|[0-9]+[.．、]|[一二三四五六七八九十]+[、.．])", p))
        if out and not is_new:
            out[-1] = normalize_sentence(out[-1] + p)
        else:
            out.append(normalize_sentence(p))
    return out


def math_r(text: str, style: dict[str, object]):
    r = etree.Element(q(M, "r"))
    rpr = etree.SubElement(r, q(A, "rPr"))
    rpr.set("lang", "en-US")
    rpr.set("altLang", "zh-CN")
    rpr.set("sz", str(int(style.get("size", 22)) * 100))
    if style.get("bold", True):
        rpr.set("b", "1")
    rpr.set("smtClean", "0")
    rpr.append(solid_fill(str(style.get("color", "555A64"))))
    etree.SubElement(rpr, q(A, "latin")).set("typeface", "Cambria Math")
    t = etree.SubElement(r, q(M, "t"))
    t.text = text
    return r


def math_group(tag: str, children: list) -> etree._Element:
    el = etree.Element(q(M, tag))
    for child in children:
        el.append(child)
    return el


def math_frac(num_children: list, den_children: list) -> etree._Element:
    frac = etree.Element(q(M, "f"))
    fpr = etree.SubElement(frac, q(M, "fPr"))
    etree.SubElement(fpr, q(M, "type")).set(q(M, "val"), "bar")
    num = etree.SubElement(frac, q(M, "num"))
    for child in num_children:
        num.append(child)
    den = etree.SubElement(frac, q(M, "den"))
    for child in den_children:
        den.append(child)
    return frac


def math_rad(children: list) -> etree._Element:
    rad = etree.Element(q(M, "rad"))
    radpr = etree.SubElement(rad, q(M, "radPr"))
    etree.SubElement(radpr, q(M, "degHide")).set(q(M, "val"), "1")
    etree.SubElement(rad, q(M, "deg"))
    e = etree.SubElement(rad, q(M, "e"))
    for child in children:
        e.append(child)
    return rad


def matching_index(text: str, start: int, open_ch: str, close_ch: str) -> int:
    depth = 0
    for idx in range(start, len(text)):
        ch = text[idx]
        if ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return idx
    return -1


def split_top_level_slash(text: str) -> tuple[str, str] | None:
    depth = 0
    for i, ch in enumerate(text):
        if ch in "([":
            depth += 1
        elif ch in ")]":
            depth = max(0, depth - 1)
        elif ch == "/" and depth == 0 and i > 0 and i < len(text) - 1:
            return text[:i], text[i + 1 :]
    return None


def parse_math_expr(expr: str, style: dict[str, object]) -> list:
    expr = expr.strip()
    expr = expr.replace("＋", "+").replace("－", "-").replace("（", "(").replace("）", ")")
    split = split_top_level_slash(expr)
    if split:
        left, right = split
        return [math_frac(parse_math_expr(left, style), parse_math_expr(right, style))]

    out: list = []
    i = 0
    buf: list[str] = []

    def flush() -> None:
        if buf:
            out.append(math_r("".join(buf), style))
            buf.clear()

    while i < len(expr):
        ch = expr[i]
        if ch == "√":
            flush()
            if i + 1 < len(expr) and expr[i + 1] in "([":
                open_ch = expr[i + 1]
                close_ch = ")" if open_ch == "(" else "]"
                end = matching_index(expr, i + 1, open_ch, close_ch)
                if end != -1:
                    out.append(math_rad(parse_math_expr(expr[i + 2 : end], style)))
                    i = end + 1
                    continue
            j = i + 1
            while j < len(expr) and re.match(r"[\w₀₁₂₃₄₅₆₇₈₉²³αβθπσμΔδ]+", expr[j], re.U):
                j += 1
            radicand = expr[i + 1 : j] or " "
            out.append(math_rad(parse_math_expr(radicand, style)))
            i = max(j, i + 1)
            continue
        if ch in "([":
            close_ch = ")" if ch == "(" else "]"
            end = matching_index(expr, i, ch, close_ch)
            if end != -1:
                flush()
                out.append(math_r(ch, style))
                out.extend(parse_math_expr(expr[i + 1 : end], style))
                out.append(math_r(close_ch, style))
                i = end + 1
                continue
        buf.append(ch)
        i += 1
    flush()
    return out


def make_formula_inline(expr: str, style: dict[str, object]):
    a14m = etree.Element(q(A14, "m"), nsmap={"a14": A14, "m": M, "a": A})
    omp = etree.SubElement(a14m, q(M, "oMathPara"))
    omppr = etree.SubElement(omp, q(M, "oMathParaPr"))
    etree.SubElement(omppr, q(M, "jc")).set(q(M, "val"), "left")
    om = etree.SubElement(omp, q(M, "oMath"))
    for child in parse_math_expr(expr, style):
        om.append(child)
    return a14m


def formula_needed(expr: str) -> bool:
    expr = expr.strip()
    if not FORMULA_RISK_RE.search(expr):
        return False
    return "√" in expr or bool(FORMULA_ONLY_RE.match(expr))


def make_text_para(text: str, style: dict[str, object], line_spacing: int | None = None, align: str | None = None):
    p_el = etree.Element(q(A, "p"))
    ppr = make_ppr(line_spacing=line_spacing, align=align)
    if line_spacing == OPTION_LINE_SPACING:
        add_spacing(ppr, before=0, after=OPTION_SPACE_AFTER)
    p_el.append(ppr)
    p_el.append(make_run(text, style))
    return p_el


def make_option_para(text: str, style: dict[str, object]):
    p_el = etree.Element(q(A, "p"))
    ppr = make_ppr(line_spacing=OPTION_LINE_SPACING)
    add_spacing(ppr, before=0, after=OPTION_SPACE_AFTER)
    p_el.append(ppr)
    m = OPTION_PREFIX_RE.match(text)
    if m and formula_needed(m.group(2)):
        p_el.append(make_run(m.group(1), style))
        p_el.append(make_formula_inline(m.group(2), style))
    else:
        p_el.append(make_run(text, style))
    return p_el


def dynamic_pictures(root) -> list:
    pics = []
    for pic in root.findall(".//p:pic", NS):
        name = shape_name(pic)
        b = shape_bounds(pic)
        if b is None:
            continue
        if re.match(r"rId\d+\.(?:png|jpg|jpeg)$", name, re.I):
            pics.append(pic)
    return pics


def safe_area(slide_w: int, slide_h: int) -> tuple[int, int, int, int]:
    w_in = to_in(slide_w)
    h_in = to_in(slide_h)
    left = emu(0.52 if w_in <= 10.5 else 0.75 if w_in <= 14 else 1.15)
    right = slide_w - emu(0.52 if w_in <= 10.5 else 0.75 if w_in <= 14 else 0.85)
    top = emu(0.52 if h_in <= 6 else 0.65)
    bottom = slide_h - emu(0.35 if h_in <= 6 else 0.55)
    return left, top, right, bottom


def move_inside(bounds: tuple[int, int, int, int], safe: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    x, y, w, h = bounds
    left, top, right, bottom = safe
    if w > right - left:
        scale = (right - left) / w
        w = int(w * scale)
        h = int(h * scale)
    if h > bottom - top:
        scale = (bottom - top) / h
        w = int(w * scale)
        h = int(h * scale)
    x = max(left, min(x, right - w))
    y = max(top, min(y, bottom - h))
    return x, y, w, h


def center_wide_images(root, slide_w: int, slide_h: int, safe: tuple[int, int, int, int]) -> int:
    changed = 0
    left, top, right, bottom = safe
    for pic in dynamic_pictures(root):
        b = shape_bounds(pic)
        if b is None:
            continue
        x, y, w, h = b
        is_wide = w > slide_w * 0.27 and h > slide_h * 0.09
        nx, ny, nw, nh = move_inside((x, y, w, h), safe)
        if is_wide:
            nx = int(left + (right - left - nw) / 2)
        if (nx, ny, nw, nh) != (x, y, w, h):
            set_bounds(pic, nx, ny, nw, nh)
            changed += 1
    return changed


def repair_text_layout(root, slide_w: int, slide_h: int) -> dict[str, int]:
    stats = {"stems": 0, "options": 0, "bodies": 0}
    safe = safe_area(slide_w, slide_h)
    left, top, right, bottom = safe
    stem_bottom = None
    pics = dynamic_pictures(root)
    right_pic_left = None
    for pic in pics:
        b = shape_bounds(pic)
        if b and b[0] > slide_w * 0.45:
            right_pic_left = b[0] if right_pic_left is None else min(right_pic_left, b[0])

    for sp in root.findall(".//p:sp", NS):
        name = shape_name(sp)
        bounds = shape_bounds(sp)
        if bounds is None:
            continue
        x, y, w, h = bounds
        paras = text_paragraphs(sp)
        if not paras:
            continue
        style = first_run_style(sp)

        if QUESTION_NAME_RE.search(name):
            text = normalize_sentence("".join(paras))
            nx = max(x, left)
            nw = max(w, right - nx)
            nh = max(h, estimated_text_height(text, nw, int(style.get("size", 22)), 1.24))
            replace_paragraphs(sp, [make_text_para(text, style, line_spacing=120000)])
            set_bounds(sp, nx, y, min(nw, right - nx), nh)
            stem_bottom = y + nh
            stats["stems"] += 1

    for sp in root.findall(".//p:sp", NS):
        name = shape_name(sp)
        bounds = shape_bounds(sp)
        if bounds is None:
            continue
        x, y, w, h = bounds
        paras = text_paragraphs(sp)
        if not paras:
            continue
        style = first_run_style(sp)

        if OPTION_NAME_RE.search(name):
            options = normalize_options(paras)
            ny = y
            if stem_bottom is not None:
                ny = max(ny, stem_bottom + emu(0.22))
            nx = max(x, left)
            max_right = right
            if right_pic_left and right_pic_left - nx > slide_w * 0.35:
                max_right = min(max_right, right_pic_left - emu(0.2))
            nw = max(emu(2.2), max_right - nx)
            nh = max(h, bottom - ny)
            replace_paragraphs(sp, [make_option_para(opt, style) for opt in options])
            set_bounds(sp, nx, ny, nw, max(emu(1.1), nh))
            stats["options"] += 1
        elif CONTENT_NAME_RE.search(name) and not LABEL_NAME_RE.search(name):
            normalized = normalize_bullets(paras)
            if normalized != paras and normalized:
                replace_paragraphs(sp, [make_text_para(p, style, line_spacing=125000) for p in normalized])
                stats["bodies"] += 1

    return stats


def repair_pptx(src: Path, dst: Path) -> dict[str, int]:
    with zipfile.ZipFile(src, "r") as zf:
        entries = {name: zf.read(name) for name in zf.namelist()}

    slide_w, slide_h = slide_size(entries)
    slide_names = sorted(
        [name for name in entries if re.match(r"ppt/slides/slide\d+\.xml$", name)],
        key=lambda name: int(re.search(r"slide(\d+)\.xml$", name).group(1)),
    )
    totals = {"slides": len(slide_names), "stems": 0, "options": 0, "bodies": 0, "images": 0}

    for name in slide_names:
        root = parse(entries[name])
        text_stats = repair_text_layout(root, slide_w, slide_h)
        for key in ("stems", "options", "bodies"):
            totals[key] += text_stats[key]
        totals["images"] += center_wide_images(root, slide_w, slide_h, safe_area(slide_w, slide_h))
        entries[name] = xml_bytes(root)

    dst.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as tmp:
        tmp_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as out:
            for name, data in entries.items():
                out.writestr(name, data)
        shutil.move(str(tmp_path), dst)
    finally:
        if tmp_path.exists():
            tmp_path.unlink()
    return totals


def main() -> int:
    parser = argparse.ArgumentParser(description="Fix generated Gaowu PPTX layout by root-cause rules.")
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    stats = repair_pptx(args.input, args.output)
    print(f"fixed {args.input.name} -> {args.output}")
    print(stats)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
