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
OPTION_SPACE_AFTER = 0

QUESTION_NAME_RE = re.compile(r"(?:Problem Stem|stem)", re.I)
OPTION_NAME_RE = re.compile(r"(?:Options|options)", re.I)
CONTENT_NAME_RE = re.compile(r"(?:Content Body|Summary Body|Text)", re.I)
LABEL_NAME_RE = re.compile(r"(?:KEY POINT|label|Option Label|Freeform)", re.I)
OPTION_PREFIX_RE = re.compile(r"^([A-D][．.]\s*)(.+)$")
FORMULA_ONLY_RE = re.compile(r"^[\sA-Za-z0-9₀₁₂₃₄₅₆₇₈₉ₐᵦαβθπσμΔδFNGQklrabvxyzXYZRLO＋+\-*/=<>≤≥（）()\[\]{}√²³^·.,，]+$")
FORMULA_RISK_RE = re.compile(r"(√|/|²|³|₀|₁|₂|₃|₄|₅|₆|₇|₈|₉|≤|≥|[A-Za-z]\s*=)")
INLINE_FORMULA_RE = re.compile(
    r"(?:[A-Za-z][A-Za-z]?\s*=\s*[A-Za-z0-9μ₀₁₂₃₄₅₆₇₈₉ₐᵦ²³]+(?:[+\-*/][A-Za-z0-9μ₀₁₂₃₄₅₆₇₈₉ₐᵦ²³]+)*)"
    r"|(?:[A-Za-z][₀₁₂₃₄₅₆₇₈₉ₐᵦ]+)"
    r"|(?:F[0-9₀₁₂₃₄₅₆₇₈₉])"
    r"|(?:[0-9]+≤[A-Za-z]≤[0-9]*N?)"
)


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
    body.set("lIns", "0")
    body.set("rIns", "0")
    body.set("tIns", "0")
    body.set("bIns", "0")
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
    # WPS/PowerPoint Chinese text wraps much closer to 1 em per Han
    # character than the earlier optimistic model. A conservative estimate is
    # required because text overflow is the highest-priority failure mode.
    chars_per_line = max(4.0, width_in * (72 / max(pt, 12)) / 0.95)
    paragraphs = [line for line in text.splitlines() if line.strip()] or [text]
    lines = sum(max(1, math.ceil(text_units(line) / chars_per_line)) for line in paragraphs)
    paragraph_gap = max(0, len(paragraphs) - 1) * 0.04
    return emu(lines * (pt / 72) * line_spacing + paragraph_gap + 0.14)


def estimated_single_line_width(text: str, pt: int) -> int:
    # This is deliberately a little generous because WPS uses text-box
    # insets and mixed CJK/Latin metrics that otherwise cause surprise wraps.
    return emu(text_units(text) * (pt / 72) * 0.74 + 0.22)


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


def option_body_text(option: str) -> str:
    match = OPTION_PREFIX_RE.match(option.strip())
    return match.group(2).strip() if match else option.strip()


def is_short_text_option_image_page(options: list[str], pics: list) -> bool:
    if len(pics) != 1 or len(options) != 4:
        return False
    bodies = [option_body_text(opt) for opt in options]
    if not all(1 <= len(body) <= 4 for body in bodies):
        return False
    b = shape_bounds(pics[0])
    if b is None or b[3] <= 0:
        return False
    return b[2] / b[3] >= 1.8


def make_text_para(text: str, style: dict[str, object], line_spacing: int | None = None, align: str | None = None):
    p_el = etree.Element(q(A, "p"))
    ppr = make_ppr(line_spacing=line_spacing, align=align)
    if line_spacing == OPTION_LINE_SPACING:
        add_spacing(ppr, before=0, after=OPTION_SPACE_AFTER)
    p_el.append(ppr)
    append_mixed_runs(p_el, text, style)
    return p_el


def append_mixed_runs(p_el, text: str, style: dict[str, object]) -> None:
    pos = 0
    for match in INLINE_FORMULA_RE.finditer(text):
        if match.start() > pos:
            p_el.append(make_run(text[pos : match.start()], style))
        p_el.append(make_formula_inline(match.group(0), style))
        pos = match.end()
    if pos < len(text):
        p_el.append(make_run(text[pos:], style))


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
        append_mixed_runs(p_el, text, style)
    return p_el


def option_label_shapes(root) -> list:
    labels = []
    for sp in root.findall(".//p:sp", NS):
        name = shape_name(sp)
        text = text_content(sp, "").strip()
        if "Option Label" in name and re.match(r"^[A-D]$", text):
            labels.append(sp)
    return labels


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
    if w_in <= 10.5:
        # The parchment/Egypt-style templates have an inner ornamental frame.
        # Content must stay inside that frame, not merely inside the slide.
        left = emu(0.62)
        right = slide_w - emu(0.55)
        top = emu(0.62)
        bottom = slide_h - emu(0.78)
    else:
        left = emu(0.75 if w_in <= 14 else 1.55)
        right = slide_w - emu(0.75 if w_in <= 14 else 1.25)
        top = emu(0.65 if w_in <= 14 else 0.9)
        bottom = slide_h - emu(0.62 if w_in <= 14 else 0.9)
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


def fit_into_box(bounds: tuple[int, int, int, int], box: tuple[int, int, int, int]) -> tuple[int, int, int, int]:
    _x, _y, w, h = bounds
    bx, by, bw, bh = box
    if w <= 0 or h <= 0 or bw <= 0 or bh <= 0:
        return bounds
    scale = min(bw / w, bh / h)
    nw = int(w * scale)
    nh = int(h * scale)
    return bx + int((bw - nw) / 2), by + int((bh - nh) / 2), nw, nh


def has_options_shape(root) -> bool:
    return any(
        OPTION_NAME_RE.search(shape_name(sp)) and text_paragraphs(sp)
        for sp in root.findall(".//p:sp", NS)
    )


def find_first_text_shape(root, pattern: re.Pattern[str]):
    for sp in root.findall(".//p:sp", NS):
        if pattern.search(shape_name(sp)) and text_paragraphs(sp):
            return sp
    return None


def layout_question_media(root, slide_w: int, slide_h: int, safe: tuple[int, int, int, int]) -> int:
    """Make stem/options/images mutually exclusive on question slides."""

    changed = 0
    stem = find_first_text_shape(root, QUESTION_NAME_RE)
    options = find_first_text_shape(root, OPTION_NAME_RE)
    pics = dynamic_pictures(root)
    if stem is None or options is None or not pics:
        return 0

    left, _top, right, bottom = safe
    gap = emu(0.3 if to_in(slide_w) <= 10.5 else 0.42)
    stem_bounds = shape_bounds(stem)
    opt_bounds = shape_bounds(options)
    if stem_bounds is None or opt_bounds is None:
        return 0

    sx, sy, sw, sh = stem_bounds
    if sx < left or sx + sw > right:
        set_bounds(stem, max(sx, left), sy, right - max(sx, left), sh)
        changed += 1
    content_top = max(opt_bounds[1], sy + sh + emu(0.22))
    content_h = max(emu(1.4), bottom - content_top)
    safe_w = right - left

    # More than two diagrams with option labels are usually image-option pages.
    # Those should remain as a centered group rather than being forced into a
    # right-side figure column.
    if len(pics) > 2:
        xs, ys, rights, bottoms = [], [], [], []
        for pic in pics:
            b = shape_bounds(pic)
            if b is None:
                continue
            xs.append(b[0])
            ys.append(b[1])
            rights.append(b[0] + b[2])
            bottoms.append(b[1] + b[3])
        if not xs:
            return changed
        group_x, group_y = min(xs), min(ys)
        group_w, group_h = max(rights) - group_x, max(bottoms) - group_y
        max_w = safe_w
        max_h = content_h
        scale = min(1.0, max_w / max(1, group_w), max_h / max(1, group_h))
        target_x = left + int((safe_w - int(group_w * scale)) / 2)
        target_y = content_top + int((content_h - int(group_h * scale)) / 2)
        dx = target_x - int(group_x * scale)
        dy = target_y - int(group_y * scale)
        for pic in pics:
            b = shape_bounds(pic)
            if b is None:
                continue
            nx = int(b[0] * scale) + dx
            ny = int(b[1] * scale) + dy
            nw = int(b[2] * scale)
            nh = int(b[3] * scale)
            set_bounds(pic, nx, ny, nw, nh)
            changed += 1
        return changed

    option_items = normalize_options(text_paragraphs(options))
    if is_short_text_option_image_page(option_items, pics):
        opt_style = first_run_style(options)
        opt_w = emu(1.28 if to_in(slide_w) <= 10.5 else 1.6)
        set_bounds(options, left, content_top, opt_w, content_h)
        replace_paragraphs(options, [make_option_para(opt, opt_style) for opt in option_items])
        b = shape_bounds(pics[0])
        if b is not None:
            fig_x = left + opt_w + gap
            fig_w = right - fig_x
            nb = fit_into_box(b, (fig_x, content_top, fig_w, content_h))
            set_bounds(pics[0], *nb)
        return changed + 2

    opt_ratio = 0.72 if to_in(slide_w) <= 10.5 else 0.6
    min_opt_w = emu(6.15 if to_in(slide_w) <= 10.5 else 7.6 if to_in(slide_w) >= 18 else 6.65)
    max_opt_w = safe_w - gap - emu(1.45 if to_in(slide_w) <= 10.5 else 3.1 if to_in(slide_w) >= 18 else 2.6)
    opt_w = max(emu(2.3), min(max_opt_w, max(min_opt_w, int(safe_w * opt_ratio))))
    opt_text = "\n".join(normalize_options(text_paragraphs(options)))
    opt_style = first_run_style(options)
    opt_size = int(opt_style.get("size", 22))
    no_wrap_w = max(
        [estimated_single_line_width(opt, opt_size) for opt in option_items] or [0]
    )
    if no_wrap_w <= max_opt_w:
        opt_w = max(opt_w, no_wrap_w)
    needed_h = estimated_text_height(opt_text, opt_w, opt_size, 1.5)
    if needed_h > content_h:
        min_fig_w = emu(1.45 if to_in(slide_w) <= 10.5 else 3.1 if to_in(slide_w) >= 18 else 2.6)
        max_by_image = safe_w - gap - min_fig_w
        target_w = int(opt_w * min(1.55, needed_h / max(1, content_h))) + emu(0.2)
        opt_w = min(max_by_image, max(opt_w, target_w, min(no_wrap_w, max_by_image)))
        needed_h = estimated_text_height(opt_text, opt_w, opt_size, 1.5)
    if needed_h > content_h * 1.05:
        shrink = min(4, math.ceil((needed_h / max(1, content_h) - 1) * 4))
        opt_style["size"] = max(16, opt_size - shrink)
    fig_x = left + opt_w + gap
    fig_w = max(emu(1.5), right - fig_x)
    set_bounds(options, left, content_top, opt_w, content_h)
    replace_paragraphs(options, [make_option_para(opt, opt_style) for opt in option_items])
    changed += 1

    if len(pics) == 1:
        b = shape_bounds(pics[0])
        if b is not None:
            nb = fit_into_box(b, (fig_x, content_top, fig_w, content_h))
            set_bounds(pics[0], *nb)
            changed += 1
    else:
        each_h = max(emu(1.0), int((content_h - gap) / 2))
        for idx, pic in enumerate(pics[:2]):
            b = shape_bounds(pic)
            if b is None:
                continue
            box_y = content_top + idx * (each_h + gap)
            nb = fit_into_box(b, (fig_x, box_y, fig_w, each_h))
            set_bounds(pic, *nb)
            changed += 1

    return changed


def layout_image_option_slide(root, slide_w: int, slide_h: int, safe: tuple[int, int, int, int]) -> int:
    """Reflow question pages whose A/B/C/D choices are diagrams."""

    stem = find_first_text_shape(root, QUESTION_NAME_RE)
    if stem is None or has_options_shape(root):
        return 0
    labels = option_label_shapes(root)
    pics = dynamic_pictures(root)
    if len(labels) < 2 or len(pics) < 3:
        return 0

    stem_bounds = shape_bounds(stem)
    if stem_bounds is None:
        return 0
    left, _top, right, bottom = safe
    gap = emu(0.24 if to_in(slide_w) <= 10.5 else 0.35)
    content_top = stem_bounds[1] + stem_bounds[3] + emu(0.28)
    content_h = max(emu(1.3), bottom - content_top)
    changed = 0

    option_pics = pics
    main_pic = None
    if len(pics) == len(labels) + 1:
        ordered = sorted(pics, key=lambda p: (shape_bounds(p) or (0, 0, 0, 0))[1])
        main_pic = ordered[0]
        option_pics = [p for p in pics if p is not main_pic]
        main_h = min(emu(1.15), max(emu(0.8), int(content_h * 0.32)))
        main_box = (right - emu(2.2), content_top, emu(2.0), main_h)
        b = shape_bounds(main_pic)
        if b is not None:
            set_bounds(main_pic, *fit_into_box(b, main_box))
            changed += 1
        content_top += main_h + gap
        content_h = max(emu(1.0), bottom - content_top)

    # If labels already form two rows, keep a 2x2 rhythm; otherwise use one
    # row for compact option images.
    label_rows = sorted({round((shape_bounds(label) or (0, 0, 0, 0))[1] / emu(0.35)) for label in labels})
    rows = 2 if len(label_rows) >= 2 else 1
    if len(option_pics) > 4:
        rows = 2
    cols = max(1, math.ceil(len(option_pics) / rows))
    cell_w = (right - left) / cols
    cell_h = content_h / rows

    sorted_pics = sorted(option_pics, key=lambda p: ((shape_bounds(p) or (0, 0, 0, 0))[1], (shape_bounds(p) or (0, 0, 0, 0))[0]))
    sorted_labels = sorted(labels, key=lambda sp: text_content(sp, ""))
    label_style = first_run_style(stem)
    label_style["color"] = str(label_style.get("color", "5A2A1B"))
    label_style["bold"] = True
    for idx, pic in enumerate(sorted_pics):
        row = idx // cols
        col = idx % cols
        cell_x = int(left + col * cell_w)
        cell_y = int(content_top + row * cell_h)
        label_w = emu(0.32)
        image_box = (
            cell_x + label_w + emu(0.08),
            cell_y,
            int(cell_w) - label_w - emu(0.12),
            max(emu(0.6), int(cell_h) - emu(0.28)),
        )
        b = shape_bounds(pic)
        placed = None
        if b is not None:
            placed = fit_into_box(b, image_box)
            set_bounds(pic, *placed)
            changed += 1
        if idx < len(sorted_labels):
            label = sorted_labels[idx]
            lb = shape_bounds(label)
            if lb is not None:
                label_text = text_content(label, "")
                label_x = cell_x + emu(0.05)
                label_y = cell_y + int(cell_h) - emu(0.38)
                if placed is not None:
                    label_x = max(left, placed[0] - lb[2] - emu(0.08))
                    label_y = min(bottom - lb[3], placed[1] + placed[3] - lb[3] - emu(0.02))
                set_bounds(label, label_x, label_y, lb[2], lb[3])
                replace_paragraphs(label, [make_text_para(label_text, label_style, align="ctr")])
                changed += 1

    return changed


def center_wide_images(root, slide_w: int, slide_h: int, safe: tuple[int, int, int, int]) -> int:
    changed = 0
    if has_options_shape(root) or len(option_label_shapes(root)) >= 2:
        return changed
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
            should_rewrite = normalized != paras or any(INLINE_FORMULA_RE.search(p) for p in (normalized or paras))
            if should_rewrite and normalized:
                replace_paragraphs(sp, [make_text_para(p, style, line_spacing=125000) for p in normalized])
                stats["bodies"] += 1
            nx = max(x, left)
            ny = max(y, top)
            nw = min(w, right - nx)
            nh = min(max(h, estimated_text_height("\n".join(normalized or paras), nw, int(style.get("size", 22)), 1.25)), bottom - ny)
            if nw > 0 and nh > 0 and (nx, ny, nw, nh) != (x, y, w, h):
                set_bounds(sp, nx, ny, nw, nh)
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
        safe = safe_area(slide_w, slide_h)
        totals["images"] += layout_question_media(root, slide_w, slide_h, safe)
        totals["images"] += layout_image_option_slide(root, slide_w, slide_h, safe)
        totals["images"] += center_wide_images(root, slide_w, slide_h, safe)
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
