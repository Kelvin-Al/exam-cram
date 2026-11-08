#!/usr/bin/env python3
"""Build a dense two-sided A4 cheat sheet from `cheatsheet.md`.

The Markdown dialect is intentionally small and print-oriented:

    ---
    course: PHYS 101 — Mechanics
    exam: Final
    date: 2026-12-20
    ---

    <!-- SIDE 1 -->
    # CORE + FORMULAS
    ## ★★★ 1. Kinematics · weight 18% · mastery 45%
    - **Definition**: ...
    $$ v = u + at \\quad x = ut + \\tfrac12 at^2 $$
    > TRAP: ...
    | Quantity | Unit | Note |
    |---|---|---|
    | Work | J | scalar |

    <!-- SIDE 2 -->
    # WEAK SPOTS + TRAPS
    ...

Build:
    python build_cheatsheet.py --md cheatsheet.md --out cheatsheet.pdf \
        --docx cheatsheet.docx --state state.json --json report.json

The script converts DOCX -> PDF with the bundled LibreOffice Kit, renders the
PDF to PNG for inspection, and shrinks the typography until the document is
exactly two A4 pages (or reports an overflow that needs trimming).
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Mm, Pt, RGBColor

# --------------------------------------------------------------------------- #
# design tokens (overridable with environment variables)
# --------------------------------------------------------------------------- #
INK = "1A1A1A"
MUTED = "5A6270"
NAVY = "17324D"
GOLD = "FFD566"
PRIORITY = {3: ("B3261E", "FCE8E6"), 2: ("B26A00", "FFF1D6"), 1: ("1F5C99", "E7F0F8")}
CALLOUT = {
    "TRAP": ("B3261E", "FDECEA"),
    "TIP": ("1F6B45", "E8F5EE"),
    "EXAM": ("1F5C99", "E8F0F8"),
    "MEMORIZE": ("6A3D9A", "F2EAF8"),
    "NOTE": ("5A6270", "F1F3F5"),
}
CALLOUT_ALIASES = {
    "陷阱": "TRAP", "易错": "TRAP", "错误": "TRAP",
    "提示": "TIP", "技巧": "TIP", "方法": "TIP",
    "考试": "EXAM", "考点": "EXAM", "得分": "EXAM",
    "必背": "MEMORIZE", "记忆": "MEMORIZE", "背下来": "MEMORIZE",
    "注意": "NOTE", "备注": "NOTE",
}
FONT_BODY = os.environ.get("EXAM_CRAM_BODY_FONT", "Arial")
FONT_CJK = os.environ.get("EXAM_CRAM_CJK_FONT", "Microsoft YaHei")
FONT_MATH = os.environ.get("EXAM_CRAM_MATH_FONT", "Cambria Math")
FONT_CODE = os.environ.get("EXAM_CRAM_CODE_FONT", "Consolas")

SCALE = 1.0
PAGE_W, PAGE_H = 210.0, 297.0
MARGIN = 5.0
COLS, GUTTER = 3, 4.0
BODY_SIZE = 6.9


def pt(value: float):
    return Pt(value * SCALE)


# --------------------------------------------------------------------------- #
# low-level DOCX helpers
# --------------------------------------------------------------------------- #
def set_run_fonts(run, ascii_font=FONT_BODY, cjk_font=FONT_CJK):
    run.font.name = ascii_font
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    rFonts.set(qn("w:ascii"), ascii_font)
    rFonts.set(qn("w:hAnsi"), ascii_font)
    rFonts.set(qn("w:eastAsia"), cjk_font)
    rFonts.set(qn("w:cs"), ascii_font)


def para_shade(p, fill):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def para_border(p, side, color, sz=12, space=2):
    pPr = p._p.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    el = OxmlElement(f"w:{side}")
    el.set(qn("w:val"), "single")
    el.set(qn("w:sz"), str(sz))
    el.set(qn("w:space"), str(space))
    el.set(qn("w:color"), color)
    pBdr.append(el)


def set_cols(section, num, spacing_mm):
    sectPr = section._sectPr
    cols = sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        sectPr.append(cols)
    cols.set(qn("w:num"), str(num))
    cols.set(qn("w:space"), str(int(spacing_mm * 56.7)))


def tight(p, before=0.0, after=0.4, line=0.95):
    pf = p.paragraph_format
    pf.space_before = pt(before)
    pf.space_after = pt(after)
    pf.line_spacing = line
    pf.widow_control = False


def shrink_break_paragraph(doc):
    """Make the paragraph that carries a section break as short as possible."""
    if not doc.paragraphs:
        return
    p = doc.paragraphs[-1]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    pPr = p._p.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), "2")
        rPr.append(el)


# --------------------------------------------------------------------------- #
# LaTeX-lite and inline Markdown
# --------------------------------------------------------------------------- #
GREEK = {
    r"\alpha": "α", r"\beta": "β", r"\gamma": "γ", r"\delta": "δ", r"\epsilon": "ε",
    r"\zeta": "ζ", r"\eta": "η", r"\theta": "θ", r"\lambda": "λ", r"\mu": "μ",
    r"\pi": "π", r"\rho": "ρ", r"\sigma": "σ", r"\tau": "τ", r"\phi": "φ",
    r"\omega": "ω", r"\Delta": "Δ", r"\Omega": "Ω", r"\Sigma": "Σ", r"\Phi": "Φ",
    r"\infty": "∞", r"\partial": "∂", r"\nabla": "∇", r"\sum": "Σ", r"\int": "∫",
    r"\times": "×", r"\cdot": "·", r"\pm": "±", r"\mp": "∓", r"\le": "≤",
    r"\leq": "≤", r"\ge": "≥", r"\geq": "≥", r"\neq": "≠", r"\approx": "≈",
    r"\equiv": "≡", r"\propto": "∝", r"\rightarrow": "→", r"\to": "→",
    r"\Rightarrow": "⇒", r"\leftrightarrow": "↔",
}
PROTECT_RE = re.compile(r"[\^_]\{[^{}]*\}")


def latex_light(s: str) -> str:
    for k, v in sorted(GREEK.items(), key=lambda kv: -len(kv[0])):
        s = s.replace(k, v)
    s = s.replace(r"\quad", "   ").replace(r"\qquad", "      ")
    s = s.replace(r"\,", " ").replace(r"\;", " ").replace(r"\!", "")
    for fn in ("sin", "cos", "tan", "log", "ln", "max", "min", "exp", "det"):
        s = s.replace("\\" + fn, fn)
    s = re.sub(r"\\text\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\mathrm\{([^{}]*)\}", r"\1", s)
    s = re.sub(r"\\tfrac\s*12", "½", s)
    s = re.sub(r"\\frac\s*12", "½", s)
    for _ in range(3):
        s = re.sub(r"\\[td]?frac\{([^{}]*)\}\{([^{}]*)\}", r"(\1)/(\2)", s)
        s = re.sub(r"\\sqrt\{([^{}]*)\}", r"√(\1)", s)
    s = s.replace("\\left", "").replace("\\right", "")
    # protect real super/subscripts, drop decorative braces, restore
    saved = []

    def _protect(m):
        saved.append(m.group(0))
        return f"\x00{len(saved) - 1}\x00"

    s = PROTECT_RE.sub(_protect, s)
    s = s.replace("{", "").replace("}", "")
    s = re.sub(r"\^([A-Za-z0-9+\-]{1,4})(?![A-Za-z0-9])", r"^{\1}", s)
    s = re.sub(r"_([A-Za-z0-9]{1,4})(?![A-Za-z0-9])", r"_{\1}", s)
    for i, token in enumerate(saved):
        s = s.replace(f"\x00{i}\x00", token)
    return s


TOKEN_RE = re.compile(
    r"(\*\*.+?\*\*|\*.+?\*|==.+?==|`.+?`|\^\{[^{}]*\}|_\{[^{}]*\}|\$[^$]+\$)"
)


def add_inline(p, text, size=BODY_SIZE, color=INK, bold=False, italic=False):
    text = latex_light(text).replace("$", "")
    pos = 0
    for m in TOKEN_RE.finditer(text):
        if m.start() > pos:
            r = p.add_run(text[pos:m.start()])
            r.font.size = pt(size)
            r.font.color.rgb = RGBColor.from_string(color)
            r.bold = bold
            r.italic = italic
            set_run_fonts(r)
        tok = m.group(0)
        code = False
        if tok.startswith("**"):
            r = p.add_run(tok[2:-2]); r.bold = True
        elif tok.startswith("=="):
            r = p.add_run(tok[2:-2])
            rPr = r._element.get_or_add_rPr()
            shd = OxmlElement("w:shd")
            shd.set(qn("w:val"), "clear"); shd.set(qn("w:color"), "auto")
            shd.set(qn("w:fill"), "FFF3B0")
            rPr.append(shd)
        elif tok.startswith("`"):
            code = True
            r = p.add_run(tok[1:-1])
            r.font.size = pt(max(5.4, size - 0.4))
            set_run_fonts(r, FONT_CODE, FONT_CJK)
            r.font.color.rgb = RGBColor.from_string("8B1A1A")
        elif tok.startswith("^{"):
            r = p.add_run(tok[2:-1]); r.font.superscript = True
        elif tok.startswith("_{"):
            r = p.add_run(tok[2:-1]); r.font.subscript = True
        elif tok.startswith("$"):
            r = p.add_run(tok[1:-1]); r.italic = True
            set_run_fonts(r, FONT_MATH, FONT_CJK)
        else:
            r = p.add_run(tok[1:-1]); r.italic = True
        if not code:
            r.font.size = pt(size)
            r.font.color.rgb = RGBColor.from_string(color)
            set_run_fonts(r)
        pos = m.end()
    if pos < len(text):
        r = p.add_run(text[pos:])
        r.font.size = pt(size)
        r.font.color.rgb = RGBColor.from_string(color)
        r.bold = bold
        r.italic = italic
        set_run_fonts(r)


# --------------------------------------------------------------------------- #
# Markdown model + parser
# --------------------------------------------------------------------------- #
@dataclass
class Block:
    kind: str
    text: str = ""
    items: list = field(default_factory=list)
    rows: list = field(default_factory=list)


@dataclass
class Section:
    title: str
    priority: int
    meta: str
    blocks: list = field(default_factory=list)


@dataclass
class Side:
    title: str
    sections: list = field(default_factory=list)
    column_breaks: set = field(default_factory=set)


def parse_md(md: str):
    md = md.lstrip("\ufeff")
    lines = md.splitlines()
    meta = {}
    i = 0
    while i < len(lines) and not lines[i].strip():
        i += 1
    if i < len(lines) and lines[i].strip() == "---":
        i += 1
        while i < len(lines) and lines[i].strip() != "---":
            if ":" in lines[i]:
                k, v = lines[i].split(":", 1)
                meta[k.strip().lower()] = v.strip()
            i += 1
        i += 1
    sides: list[Side] = []
    cur_side = None
    cur_sec = None
    while i < len(lines):
        raw = lines[i].rstrip("\n")
        line = raw.strip()
        if re.match(r"<!--\s*SIDE\s*[12]\s*-->", line, re.I):
            cur_side = Side(title="")
            sides.append(cur_side)
            cur_sec = None
            i += 1
            continue
        if re.match(r"<!--\s*COLUMN\s*-->", line, re.I):
            if cur_side is not None:
                cur_side.column_breaks.add(len(cur_side.sections))
            i += 1
            continue
        if cur_side is None:
            i += 1
            continue
        if line.startswith("# "):
            cur_side.title = line[2:].strip()
        elif line.startswith("## "):
            t = line[3:].strip()
            stars = 0
            while t.startswith("★"):
                stars += 1
                t = t[1:].strip()
            meta_bits = []
            while True:
                m = re.search(r"\s*·\s*([^·]+)$", t)
                if not m:
                    break
                meta_bits.insert(0, m.group(1).strip())
                t = t[:m.start()].strip()
            cur_sec = Section(
                title=t,
                priority=max(1, min(3, stars or 2)),
                meta=" · ".join(meta_bits),
            )
            cur_side.sections.append(cur_sec)
        elif line.startswith("### "):
            if cur_sec:
                cur_sec.blocks.append(Block("sub", line[4:].strip()))
        elif line.startswith("$$"):
            buf = line[2:]
            if buf.endswith("$$"):
                buf = buf[:-2]
            else:
                i += 1
                parts = [buf]
                while i < len(lines) and not lines[i].strip().endswith("$$"):
                    parts.append(lines[i].strip())
                    i += 1
                if i < len(lines):
                    parts.append(lines[i].strip()[:-2])
                buf = " ".join(parts)
            if cur_sec:
                cur_sec.blocks.append(Block("formula", buf.strip()))
        elif line.startswith(">"):
            if cur_sec:
                cur_sec.blocks.append(Block("callout", line[1:].strip()))
        elif line.startswith("|") and line.endswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
                    rows.append(cells)
                i += 1
            if cur_sec and rows:
                cur_sec.blocks.append(Block("table", rows=rows))
            continue
        elif re.match(r"^\s*[-*]\s+", raw):
            if cur_sec:
                indent = len(raw) - len(raw.lstrip())
                cur_sec.blocks.append(
                    Block("bullet", re.sub(r"^\s*[-*]\s+", "", raw), items=[str(indent)])
                )
        elif re.match(r"^\s*\d+\.\s+", raw):
            if cur_sec:
                cur_sec.blocks.append(Block("numbered", re.sub(r"^\s*\d+\.\s+", "", raw)))
        elif line and not line.startswith("<!--"):
            if cur_sec:
                cur_sec.blocks.append(Block("para", line))
        i += 1
    return meta, sides


# --------------------------------------------------------------------------- #
# renderers
# --------------------------------------------------------------------------- #
def add_header_band(doc, n, title, meta):
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tblPr = table._tbl.tblPr
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(200 * 56.7)))
    tblPr.insert(0, tblW)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "nil")
        borders.append(el)
    tblPr.append(borders)
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", 10), ("bottom", 10), ("left", 80), ("right", 80)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)
    grid = table._tbl.tblGrid
    for gc, mm in zip(grid.findall(qn("w:gridCol")), (145, 55)):
        gc.set(qn("w:w"), str(int(mm * 56.7)))
    tr = table.rows[0]
    tr.height = Mm(10.4)
    left, right = tr.cells
    left.width = Mm(145)
    right.width = Mm(55)
    for cell in (left, right):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), NAVY)
        tcPr.append(shd)
    p = left.paragraphs[0]
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    r = p.add_run(title)
    r.bold = True
    r.font.size = pt(12.4)
    r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    set_run_fonts(r)
    if meta:
        r2 = p.add_run(f"   {meta}")
        r2.font.size = pt(5.7)
        r2.font.color.rgb = RGBColor(0xD6, 0xE2, 0xEE)
        set_run_fonts(r2)
    p2 = right.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    r3 = p2.add_run(f"SIDE {n}/2")
    r3.font.size = pt(7.4)
    r3.bold = True
    r3.font.color.rgb = RGBColor.from_string(GOLD)
    set_run_fonts(r3)


def add_section_heading(doc, sec):
    color, fill = PRIORITY[sec.priority]
    p = doc.add_paragraph()
    tight(p, 1.4, 0.3, 0.92)
    p.paragraph_format.keep_with_next = True
    p.paragraph_format.left_indent = Mm(1.2)
    para_shade(p, fill)
    para_border(p, "left", color, 18, 2)
    r = p.add_run("★" * sec.priority + " ")
    r.font.size = pt(6.8)
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(color)
    set_run_fonts(r)
    r = p.add_run(sec.title)
    r.font.size = pt(8.2)
    r.bold = True
    r.font.color.rgb = RGBColor.from_string(color)
    set_run_fonts(r)
    if sec.meta:
        r = p.add_run(f"  {sec.meta}")
        r.font.size = pt(5.6)
        r.font.color.rgb = RGBColor.from_string(MUTED)
        set_run_fonts(r)


def add_bullet(doc, text, indent=0):
    p = doc.add_paragraph()
    tight(p, 0, 0.1, 0.94)
    level = min(1, indent // 2)
    p.paragraph_format.left_indent = Mm(2.0 + 1.8 * level)
    p.paragraph_format.first_line_indent = Mm(-1.7)
    if text.lstrip().startswith("[ ]"):
        text = text.lstrip()[3:].strip()
        r = p.add_run("☐ ")
        r.font.size = pt(6.8)
        set_run_fonts(r)
    else:
        r = p.add_run("• " if level == 0 else "– ")
        r.font.size = pt(6.8)
        r.bold = True
        r.font.color.rgb = RGBColor.from_string("6B7280")
        set_run_fonts(r)
    add_inline(p, text)


def add_numbered(doc, text):
    p = doc.add_paragraph()
    tight(p, 0, 0.1, 0.94)
    p.paragraph_format.left_indent = Mm(2.4)
    p.paragraph_format.first_line_indent = Mm(-1.7)
    r = p.add_run("▸ ")
    r.font.size = pt(6.8)
    r.bold = True
    r.font.color.rgb = RGBColor.from_string("1F5C99")
    set_run_fonts(r)
    add_inline(p, text)


def add_sub(doc, text):
    p = doc.add_paragraph()
    tight(p, 0.8, 0.1, 0.92)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.bold = True
    r.font.size = pt(7.2)
    r.font.color.rgb = RGBColor.from_string("2B3A4A")
    set_run_fonts(r)


def add_formula(doc, text):
    p = doc.add_paragraph()
    tight(p, 0.6, 0.6, 0.92)
    p.paragraph_format.left_indent = Mm(1.0)
    p.paragraph_format.right_indent = Mm(0.4)
    para_shade(p, "EEF3F8")
    para_border(p, "left", "1F5C99", 10, 2)
    text = text.replace("$", "")
    for idx, part in enumerate(re.split(r"\\\\", text)):
        if idx:
            p.add_run().add_break()
        add_inline(p, part.strip(), size=7.1)


def add_callout(doc, text):
    m = re.match(r"^([^:：]{1,12})\s*[:：]\s*(.*)$", text, re.S)
    if m:
        raw_label = m.group(1).strip()
        label = raw_label.upper()
        label = CALLOUT_ALIASES.get(raw_label, label)
        body = m.group(2)
        if label not in CALLOUT:
            label = "NOTE"
    else:
        label, body = "NOTE", text
    color, fill = CALLOUT[label]
    display = raw_label if m and re.search(r"[\u4e00-\u9fff]", raw_label) else label
    p = doc.add_paragraph()
    tight(p, 0.6, 0.6, 0.93)
    p.paragraph_format.left_indent = Mm(1.3)
    p.paragraph_format.right_indent = Mm(0.3)
    para_shade(p, fill)
    para_border(p, "left", color, 14, 2)
    r = p.add_run(f"{display} ")
    r.bold = True
    r.font.size = pt(6.2)
    r.font.color.rgb = RGBColor.from_string(color)
    set_run_fonts(r)
    add_inline(p, body, size=6.5)


def add_table(doc, rows):
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    rows = [r + [""] * (ncols - len(r)) for r in rows]
    table = doc.add_table(rows=len(rows), cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    table.autofit = False
    tblPr = table._tbl.tblPr
    mar = OxmlElement("w:tblCellMar")
    for side, val in (("top", 0), ("bottom", 0), ("left", 20), ("right", 20)):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")
        mar.append(el)
    tblPr.append(mar)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single" if side == "insideH" else "nil")
        el.set(qn("w:sz"), "2")
        el.set(qn("w:color"), "C9D2DC")
        borders.append(el)
    tblPr.append(borders)
    section = doc.sections[-1]
    usable_mm = (section.page_width - section.left_margin - section.right_margin) / 36000
    col_mm = (usable_mm - 2 * GUTTER) / COLS
    tblW = OxmlElement("w:tblW")
    tblW.set(qn("w:type"), "dxa")
    tblW.set(qn("w:w"), str(int(col_mm * 56.7)))
    tblPr.insert(0, tblW)
    grid = table._tbl.tblGrid
    for gc in grid.findall(qn("w:gridCol")):
        gc.set(qn("w:w"), str(int(col_mm / ncols * 56.7)))
    for ri, row in enumerate(rows):
        tr = table.rows[ri]
        trPr = tr._tr.get_or_add_trPr()
        trPr.append(OxmlElement("w:cantSplit"))
        for ci, cell_text in enumerate(row):
            cell = tr.cells[ci]
            cell.width = Mm(col_mm / ncols)
            para = cell.paragraphs[0]
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            para.paragraph_format.line_spacing = 0.88
            if ri == 0:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:color"), "auto")
                shd.set(qn("w:fill"), "E8EDF3")
                tcPr.append(shd)
            add_inline(para, cell_text, size=6.2 / max(SCALE, 0.1), bold=(ri == 0))


def render_header(doc, side, meta, n):
    title = side.title or meta.get("course", "Cheat Sheet")
    band = " · ".join(x for x in (meta.get("course", ""), meta.get("exam", ""), meta.get("date", "")) if x)
    add_header_band(doc, n, title, band)


def column_break(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    r = p.add_run()
    r.font.size = Pt(1)
    r.add_break(WD_BREAK.COLUMN)


def render_body(doc, side):
    for idx, sec in enumerate(side.sections):
        if idx in side.column_breaks and idx > 0:
            column_break(doc)
        add_section_heading(doc, sec)
        for b in sec.blocks:
            if b.kind == "bullet":
                add_bullet(doc, b.text, int(b.items[0]) if b.items else 0)
            elif b.kind == "numbered":
                add_numbered(doc, b.text)
            elif b.kind == "sub":
                add_sub(doc, b.text)
            elif b.kind == "formula":
                add_formula(doc, b.text)
            elif b.kind == "callout":
                add_callout(doc, b.text)
            elif b.kind == "table":
                add_table(doc, b.rows)
            elif b.kind == "para":
                p = doc.add_paragraph()
                tight(p, 0.2, 0.3, 0.94)
                add_inline(p, b.text, size=6.6)


def setup_page(section):
    section.page_width = Mm(PAGE_W)
    section.page_height = Mm(PAGE_H)
    section.top_margin = Mm(MARGIN)
    section.bottom_margin = Mm(MARGIN)
    section.left_margin = Mm(MARGIN)
    section.right_margin = Mm(MARGIN)
    section.header_distance = Mm(0)
    section.footer_distance = Mm(0)


def _patch_unused_courier(docx_path: Path):
    """The python-docx starter template embeds an unused Courier style.

    LibreOffice reports it as a missing font even though no run uses it. Replace
    the two XML parts that carry it, rebuilding the archive so no member is
    duplicated.
    """
    import zipfile
    tmp = docx_path.with_suffix(docx_path.suffix + ".tmp")
    with zipfile.ZipFile(docx_path, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename in {"word/stylesWithEffects.xml", "word/fontTable.xml"}:
                data = data.replace(b"Courier", b"Consolas")
            zout.writestr(item, data)
    tmp.replace(docx_path)


def build_docx(md_text: str, out_docx: Path):
    meta, sides = parse_md(md_text)
    if len(sides) != 2:
        raise ValueError(
            f"expected exactly 2 sides marked with <!-- SIDE 1 --> and <!-- SIDE 2 -->; found {len(sides)}"
        )
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = FONT_BODY
    st.font.size = pt(BODY_SIZE)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT_CJK)
    st.paragraph_format.space_before = Pt(0)
    st.paragraph_format.space_after = Pt(0)
    st.paragraph_format.line_spacing = 0.95
    # The python-docx starter template carries an unused Courier style; avoid a
    # false missing-font warning from LibreOffice.
    for style in doc.styles:
        try:
            if style.font and style.font.name == "Courier":
                style.font.name = FONT_CODE
        except Exception:
            pass

    doc.core_properties.title = f"{meta.get('course', 'Cheat Sheet')} — {meta.get('exam', '')}".strip(" —")
    doc.core_properties.subject = "Two-sided A4 exam cheat sheet"
    doc.core_properties.comments = "Generated by the exam-cram skill."

    s0 = doc.sections[0]
    setup_page(s0)
    set_cols(s0, 1, 0)
    render_header(doc, sides[0], meta, 1)

    body1 = doc.add_section(WD_SECTION.CONTINUOUS)
    shrink_break_paragraph(doc)
    setup_page(body1)
    set_cols(body1, COLS, GUTTER)
    render_body(doc, sides[0])

    hdr2 = doc.add_section(WD_SECTION.NEW_PAGE)
    shrink_break_paragraph(doc)
    setup_page(hdr2)
    set_cols(hdr2, 1, 0)
    render_header(doc, sides[1], meta, 2)

    body2 = doc.add_section(WD_SECTION.CONTINUOUS)
    shrink_break_paragraph(doc)
    setup_page(body2)
    set_cols(body2, COLS, GUTTER)
    render_body(doc, sides[1])
    out_docx.parent.mkdir(parents=True, exist_ok=True)
    doc.save(out_docx)
    _patch_unused_courier(out_docx)


# --------------------------------------------------------------------------- #
# conversion / rendering
# --------------------------------------------------------------------------- #
def find_node():
    env = os.environ.get("DSH_NODE_EXECUTABLE")
    if env and Path(env).exists():
        return Path(env)
    home = Path.home()
    for pattern in (
        ".dsh/dsh-runtimes/*/dependencies/node/bin/node.exe",
        ".dsh/dsh-runtimes/*/dependencies/node/bin/node",
    ):
        hits = sorted(home.glob(pattern))
        if hits:
            return hits[0]
    which = shutil.which("node")
    if which:
        return Path(which)
    for p in (r"C:\Program Files\nodejs\node.exe", "/usr/local/bin/node", "/usr/bin/node"):
        if Path(p).exists():
            return Path(p)
    return None


def find_lo_cli():
    env = os.environ.get("DSH_OFFICE_CLI")
    if env and Path(env).exists():
        return Path(env)
    home = Path.home()
    rel = Path("app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js")
    roots = []
    if os.name == "nt":
        local = os.environ.get("LOCALAPPDATA")
        if local:
            roots.append(Path(local) / "Programs" / "DeepSeek Harness" / "resources")
        roots += [
            Path(r"C:\Program Files\DeepSeek Harness\resources"),
            Path(r"C:\Program Files (x86)\DeepSeek Harness\resources"),
        ]
    elif sys.platform == "darwin":
        roots += [Path("/Applications/DeepSeek Harness.app/Contents/Resources")]
    else:
        roots += [Path("/opt/DeepSeek Harness/resources"),
                  home / ".local/share/DeepSeek Harness/resources"]
    for root in roots:
        p = root / rel
        if p.exists():
            return p
    # last resort: search the DSH runtime tree shallowly
    for base in (home / ".dsh",):
        hits = list(base.glob("**/libreoffice-kit/lib/cli.js"))
        if hits:
            return hits[0]
    which = shutil.which("dsoffice")
    return Path(which) if which else None


def convert(docx: Path, pdf: Path, node: Path, cli: Path):
    pdf.unlink(missing_ok=True)
    res = subprocess.run(
        [str(node), str(cli), "convert", "--input", str(docx), "--output", str(pdf)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240,
    )
    if res.returncode != 0 or not pdf.exists():
        raise RuntimeError(f"LibreOffice conversion failed: {res.stderr or res.stdout}")
    head = pdf.read_bytes()[:5]
    if head != b"%PDF-":
        raise RuntimeError("conversion produced a file without a %PDF- signature")
    try:
        info = json.loads(res.stdout.strip() or "{}")
    except json.JSONDecodeError:
        info = {}
    return info


def render_pages(pdf: Path, out_dir: Path, node: Path, cli: Path):
    shutil.rmtree(out_dir, ignore_errors=True)
    res = subprocess.run(
        [str(node), str(cli), "render", "--input", str(pdf), "--output-dir", str(out_dir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=240,
    )
    if res.returncode != 0:
        raise RuntimeError(f"LibreOffice render failed: {res.stderr or res.stdout}")
    return json.loads(res.stdout)


# --------------------------------------------------------------------------- #
# coverage + fit
# --------------------------------------------------------------------------- #
def _norm_name(text: str) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", (text or "").lower()).strip()


def measure_fill(preview_dir: Path):
    """Rough vertical fill ratio per page, used to flag an underfull sheet."""
    try:
        from PIL import Image
    except Exception:
        return []
    out = []
    for p in sorted(preview_dir.glob("page-*.png")):
        try:
            im = Image.open(p).convert("L")
            w, h = im.size
            pix = im.load()
            last = 0
            for y in range(0, h, 3):
                for x in range(0, w, 6):
                    if pix[x, y] < 245:
                        last = y
                        break
            out.append({"page": p.name, "fill": round(last / h, 3)})
        except Exception:
            continue
    return out


def coverage_check(md_text: str, state_path: Path | None):
    if not state_path:
        return {"checked": False}
    state = json.loads(state_path.read_text(encoding="utf-8-sig"))
    text = md_text.lower()
    norm_text = _norm_name(md_text)
    missing = []
    for t in state.get("topics", []):
        tid = str(t.get("id", "")).strip().lower()
        name = _norm_name(str(t.get("name", "")))
        id_hit = bool(tid) and re.search(rf"(?<![a-z0-9]){re.escape(tid)}(?![a-z0-9])", text)
        name_hit = bool(name) and name in norm_text
        if not (id_hit or name_hit):
            missing.append(f"{t.get('id')} {t.get('name')}")
    return {
        "checked": True,
        "topics": len(state.get("topics", [])),
        "covered": len(state.get("topics", [])) - len(missing),
        "missing": missing,
    }


def build_once(md_text: str, docx: Path, pdf: Path, node: Path, cli: Path, preview: Path):
    build_docx(md_text, docx)
    info = convert(docx, pdf, node, cli)
    manifest = render_pages(pdf, preview, node, cli)
    return info, manifest


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--md", required=True)
    ap.add_argument("--out", required=True, help="final PDF path")
    ap.add_argument("--docx", default="", help="editable DOCX path (default: alongside PDF)")
    ap.add_argument("--state", default="", help="state.json for coverage check")
    ap.add_argument("--json", default="", help="write a build report JSON")
    ap.add_argument("--preview-dir", default="", help="PNG preview directory")
    ap.add_argument("--min-scale", type=float, default=0.88)
    ap.add_argument("--max-scale", type=float, default=1.0)
    ap.add_argument("--max-grow", type=float, default=1.50,
                    help="largest scale to try when the sheet is underfull")
    ap.add_argument("--step", type=float, default=0.03)
    ap.add_argument("--no-fit", action="store_true", help="build only at max scale")
    ap.add_argument("--strict-coverage", action="store_true",
                    help="exit 3 when a state topic is absent from the Markdown")
    args = ap.parse_args(argv)

    md_path = Path(args.md).expanduser().resolve()
    if not md_path.exists():
        raise SystemExit(f"Markdown not found: {md_path}")
    md_text = md_path.read_text(encoding="utf-8-sig")
    out_pdf = Path(args.out).expanduser().resolve()
    out_docx = Path(args.docx).expanduser().resolve() if args.docx else out_pdf.with_suffix(".docx")
    preview_dir = (Path(args.preview_dir).expanduser().resolve()
                   if args.preview_dir else out_pdf.with_name(out_pdf.stem + "-preview"))
    state_path = Path(args.state).expanduser().resolve() if args.state else None

    cov = coverage_check(md_text, state_path)
    if cov.get("missing"):
        print("coverage warning: missing topics: " + "; ".join(cov["missing"]), file=sys.stderr)

    node = find_node()
    cli = find_lo_cli()
    if not node or not cli:
        # Still produce the editable DOCX so the office skills can finish the job.
        global SCALE
        SCALE = 1.0
        build_docx(md_text, out_docx)
        report = {
            "status": "docx_only",
            "docx": str(out_docx),
            "coverage": cov,
            "reason": "node or LibreOffice Kit CLI not found",
        }
        print(json.dumps(report, ensure_ascii=False, indent=2))
        if args.json:
            Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                                       encoding="utf-8")
        raise SystemExit(2)

    tmp_root = out_pdf.parent / ".exam-cram-tmp"
    tmp_root.mkdir(parents=True, exist_ok=True)
    build_dir = tmp_root / f"build-{os.getpid()}-{int(time.time())}"
    build_dir.mkdir(parents=True, exist_ok=True)
    tmp = build_dir
    cache = {}

    def evaluate(scale_value):
        scale_value = round(scale_value, 3)
        if scale_value in cache:
            return cache[scale_value]
        global SCALE
        SCALE = scale_value
        stage_docx = tmp / f"stage-{scale_value:.2f}.docx"
        stage_pdf = tmp / f"stage-{scale_value:.2f}.pdf"
        stage_preview = tmp / f"preview-{scale_value:.2f}"
        info, manifest = build_once(md_text, stage_docx, stage_pdf, node, cli, stage_preview)
        pages = int(manifest.get("pageCount", 0))
        print(f"scale {scale_value:.2f}: {pages} page(s)")
        result = (pages, stage_pdf, info, manifest, stage_preview)
        cache[scale_value] = result
        return result

    try:
        if args.no_fit:
            chosen_scale = round(args.max_scale, 3)
            p1 = evaluate(chosen_scale)[0]
        else:
            p1 = evaluate(1.0)[0]
            if p1 == 2:
                # Underfull sheets should use larger type rather than three tiny columns.
                pmax = evaluate(args.max_grow)[0]
                if pmax == 2:
                    chosen_scale = round(args.max_grow, 3)
                else:
                    lo, hi = 1.0, round(args.max_grow, 3)
                    while hi - lo > args.step + 1e-9:
                        mid = round((lo + hi) / 2, 3)
                        if evaluate(mid)[0] == 2:
                            lo = mid
                        else:
                            hi = mid
                    chosen_scale = lo
            elif p1 > 2:
                pmin = evaluate(args.min_scale)[0]
                if pmin > 2:
                    chosen_scale = round(args.min_scale, 3)
                else:
                    lo, hi = round(args.min_scale, 3), 1.0
                    while hi - lo > args.step + 1e-9:
                        mid = round((lo + hi) / 2, 3)
                        if evaluate(mid)[0] == 2:
                            lo = mid
                        else:
                            hi = mid
                    chosen_scale = lo
            else:
                chosen_scale = 1.0
        pages, stage_pdf, info, manifest, stage_preview = evaluate(chosen_scale)
        scale = chosen_scale
        out_pdf.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(stage_pdf, out_pdf)
        # rebuild the final DOCX at the chosen scale, then copy previews
        SCALE = scale
        build_docx(md_text, out_docx)
        if preview_dir.exists():
            shutil.rmtree(preview_dir, ignore_errors=True)
        preview_dir.parent.mkdir(parents=True, exist_ok=True)
        shutil.copytree(stage_preview, preview_dir)
    finally:
        shutil.rmtree(build_dir, ignore_errors=True)

    status = "ok" if pages == 2 else ("overflow" if pages and pages > 2 else "underflow")
    fill = measure_fill(preview_dir)
    underfull = [f["page"] for f in fill if f["fill"] < 0.62]
    report = {
        "status": status,
        "pages": pages,
        "scale": scale,
        "pdf": str(out_pdf),
        "docx": str(out_docx),
        "preview_dir": str(preview_dir),
        "preview_images": [str(p) for p in sorted(preview_dir.glob("page-*.png"))],
        "page_fill": fill,
        "coverage": cov,
        "missing_fonts": info.get("missingFonts", []) if isinstance(info, dict) else [],
        "note": ("exactly 2 A4 sides" if status == "ok"
                 else "content does not fit 2 pages at the minimum scale; trim or rebalance"),
    }
    if underfull:
        report["underfull_pages"] = underfull
        report["note"] += "; some pages are underfull — add useful detail (formula bank, archetypes, traps), not filler"
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                                   encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if status == "overflow":
        raise SystemExit(4)
    if args.strict_coverage and cov.get("missing"):
        raise SystemExit(3)
    return 0


if __name__ == "__main__":
    sys.exit(main())
