# Cheat-sheet Markdown format

`build_cheatsheet.py` reads one Markdown file and produces a 3-column, two-sided A4 PDF (plus an editable DOCX). The dialect is deliberately small: if a construct is not listed here, it is rendered as plain text.

## Skeleton

```markdown
---
course: PHYS 101 — Mechanics
exam: Final
date: 2026-12-20
---

<!-- SIDE 1 -->
# CORE + FORMULAS

## ★★★ 1. Kinematics · weight 18% · mastery 45%
- **Definition**: ...
$$ v = u + at \quad x = ut + \tfrac12 at^2 $$
> TRAP: ...

<!-- SIDE 2 -->
# WEAK SPOTS + TRAPS

## ★★★ 4. Rotational dynamics · weight 20% · mastery 25%
...
```

- Frontmatter is optional but recommended. Keys: `course`, `exam`, `date`. They appear in the header band.
- Exactly two side markers are required: `<!-- SIDE 1 -->` and `<!-- SIDE 2 -->`. Case-insensitive.
- The first `# ` heading after a side marker is that side's header-band title. Use a short, all-caps label such as `CORE + FORMULAS` or `WEAK SPOTS + TRAPS`.
- A side may contain any number of `##` sections.
- Optional `<!-- COLUMN -->` marker, placed between sections, forces the next section into the next column. Use it on short sheets to avoid a page with one full column and two empty ones. Do not use more than two markers per side.
- For long content, omit the markers and let the columns flow naturally.

```markdown
<!-- SIDE 1 -->
# CORE + FORMULAS
## ★★★ 1. Kinematics ...
...
<!-- COLUMN -->
## ★★★ 2. Newton ...
...
<!-- COLUMN -->
## ★★ 3. Energy ...
```

## Section headings

```markdown
## ★★★ 3. Energy and momentum · weight 12% · mastery 30%
```

- `★` count = priority: 3 = must know (red), 2 = should know (amber), 1 = nice to know (blue). No stars defaults to 2.
- The metadata after the title is optional and separated by ` · `. Keep it short: `weight 18%`, `mastery 45%`, `exam 8 marks`. It is rendered in small grey text.
- The title should start with a number or a short name; it is not wrapped automatically, so keep it under ~70 characters.

## Bullets and sub-bullets

```markdown
- **Definition**: displacement is a vector; distance is a scalar.
- Formula: $v = u + at$, valid for constant acceleration.
  - Sub-bullet: choose a sign convention first.
- [ ] Checklist item (renders as ☐).
```

- One bullet = one line in the Markdown. Do not continue a bullet on the next line; start a new bullet.
- Indent two spaces for a sub-bullet.
- `[ ]` at the start of the text renders a checkbox.
- Avoid long paragraphs; the builder does not hyphenate and narrow columns punish long lines.

## Formulas

Inline:

```markdown
- Work–energy theorem: $W_{net} = \Delta K$.
```

Block (shaded box, left rule):

```markdown
$$ v = u + at \qquad x = ut + \tfrac12 at^2 \qquad v^2 = u^2 + 2ax $$
```

- Separate multiple equations on one line with `\qquad` or `\quad`.
- Use `\\` inside a block to force a line break.
- Supported LaTeX-lite: `\frac{a}{b}`, `\sqrt{x}`, `\alpha`…`\omega`, `\Delta`, `\Sigma`, `\infty`, `\partial`, `\nabla`, `\sum`, `\int`, `\times`, `\cdot`, `\pm`, `\le`, `\ge`, `\neq`, `\approx`, `\equiv`, `\to`, `\Rightarrow`, `\text{...}`, `\mathrm{...}`.
- `\tfrac12`/`\frac12` becomes `½`.
- Braces are optional for single-character super/subscripts (`v^2`, `f_k`), but use braces for multi-character ones (`W_{net}`, `I_{cm}`, `v_{initial}`). The converter protects braced scripts.
- `\left`/`\right` are stripped. Do not rely on `\begin{aligned}` or matrices; use plain lines.
- The math font is Cambria Math with a CJK fallback. If a symbol is missing on the target machine, replace it with a Unicode equivalent.

## Callouts

```markdown
> TRAP: average velocity is not (u+v)/2 unless acceleration is constant.
> TIP: check units before substituting.
> EXAM: draw the free-body diagram; marks are for the diagram.
> MEMORIZE: ...
> NOTE: ...
```

- Recognised labels: `TRAP` (red), `TIP` (green), `EXAM` (blue), `MEMORIZE` (purple), `NOTE` (grey). Chinese aliases work: `陷阱`/`易错` → TRAP, `提示`/`技巧` → TIP, `考试`/`考点` → EXAM, `必背`/`记忆` → MEMORIZE, `注意` → NOTE.
- A `>` block without a label renders as `NOTE`.
- Keep callouts to 1–3 lines; they are expensive in column space.

## Tables

```markdown
| Quantity | Unit | Note |
|---|---|---|
| Work | J | scalar |
| Power | W | $P = Fv$ |
| Momentum | kg·m/s | vector |
```

- The first row is the header; the second row must be a separator (`|---|---|`).
- Tables are sized to one column (~64 mm) and use a light header shading. Use **2 columns** for prose comparisons (situation → use) and 3 columns only for short numeric data. A 3-column table with long phrases becomes very tall.
- Table text is capped at about 6.2 pt so it does not balloon when the sheet is enlarged; body text may be larger.
- Tables cannot break across lines; a tall table may push content to the next column, which is fine.

## What goes on each side

### Side 1 — CORE + FORMULAS

- A top `> MEMORIZE:` or `> NOTE:` box with the "if you only remember 8 things" list.
- Highest-weight topics first. Each has: a one-line definition, the formulas with conditions, a standard method, and one trap.
- A formula bank table if the course has many formulas: formula | when to use | condition | common mistake.
- Unit/sign conventions, graph shapes, and "recognise this" cues.

### Side 2 — WEAK SPOTS + TRAPS

- Lowest-mastery / highest-risk topics first (use `mastery.json` space %).
- Discrimination tables: "X vs Y — when to use which".
- Decision trees written as indented bullets.
- Exam-day checklist (`- [ ] ...`).
- A final "last 30 minutes" box with 3–6 items.

## Space budget

`exam_cram.py mastery --write` produces `space_pct` per topic. Use it as a budget:

- 20–24%: weak, high-weight topic; give it a full section with formulas, worked pattern, and traps.
- 12–18%: normal topic; definition + 2–4 formulas + one trap.
- 6–10%: low-weight or solid topic; one line or a table row.
- Every topic gets at least one mention, even at 1–2 lines.

If the builder reports `overflow`, cut in this order:

1. repeated wording and "nice to know" history;
2. long derivations (keep the result and the condition);
3. merge two low-weight topics under one heading;
4. reduce examples to a single "pattern" line.

Do **not** cut a whole topic, and do not shrink below the builder's 0.88 scale; 6 pt is already small.

## Build and inspect

```powershell
& "<python>" "<skill-dir>/scripts/build_cheatsheet.py" `
  --md "cheatsheet.md" --out "cheatsheet.pdf" --docx "cheatsheet.docx" `
  --state "state.json" --json "build-report.json"
```

The script prints a JSON report:

| Field | Meaning |
|---|---|
| `status` | `ok`, `overflow`, `underflow`, or `docx_only` |
| `pages` | PDF page count; must be 2 |
| `scale` | typography scale used (0.88 → 1.50); short sheets are enlarged for readability |
| `coverage.missing` | state topics not found in the Markdown; fix before delivering |
| `preview_images` | PNGs to inspect with `read_image` |
| `page_fill` | rough vertical fill per page; below 0.62 is flagged |
| `underfull_pages` | pages that need more useful detail (not filler) |
| `missing_fonts` | fonts LibreOffice could not find; ignore only if the preview looks right |

The fit loop treats 1.00 as a starting point, then searches upward to `--max-grow` (default 1.50) for a short sheet, or downward to `--min-scale` (default 0.88) for a dense one. It keeps the largest scale that still produces exactly 2 pages.

## Visual QA checklist

- [ ] Two pages, both sides used; no third page.
- [ ] No text clipped at the right edge of a column or at the bottom of the page.
- [ ] No section heading stranded at the bottom of a column.
- [ ] Three columns reasonably balanced; no empty column.
- [ ] Body text legible at print size (≥6 pt); Chinese characters not boxes.
- [ ] Superscripts, subscripts, arrows, units, and minus signs correct.
- [ ] Highest-weight/weakest topics visibly get the most space.
- [ ] Header band shows course, exam, date, and `SIDE 1/2` / `SIDE 2/2`.

## Troubleshooting

- **`expected exactly 2 sides`** — add the missing `<!-- SIDE 2 -->` marker.
- **`coverage.missing`** — the check is literal (ID with word boundaries, or full name). Add the exact ID or name somewhere in the Markdown.
- **`overflow` at 0.88** — cut content; 40+ dense sections will not fit. A realistic target is 10–18 topics with 4–10 lines each. If page 1 is short but page 3 exists, the `<!-- COLUMN -->` markers are forcing content too early: remove some markers and let the columns flow.
- **A page is underfull while another page overflows** — remove `<!-- COLUMN -->` markers first; they are a balancing aid, not a requirement.
- **Table spills out of its column** — reduce to 2–3 columns or shorten cell text; the builder fixes the table width to one column.
- **Formula shows `^` or `_` literally** — use `^{...}` / `_{...}` or a single character; ensure the brace is balanced.
- **CJK looks wrong** — set `EXAM_CRAM_CJK_FONT` to an installed font (e.g. `Microsoft YaHei`, `Noto Sans CJK SC`, `PingFang SC`) and rebuild.
- **`docx_only`** — load `office-export-pdf`, convert the DOCX, then `office-render` to count pages; adjust the Markdown and repeat. On a normal DSH install the bundled LibreOffice Kit is auto-detected.
