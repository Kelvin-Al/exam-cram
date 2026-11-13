# Cheat-sheet authoring guide (LaTeX)

`build_cheatsheet.py` compiles `cheatsheet.tex` with **tectonic** and auto-fits
the font so the output is exactly **two A4 pages** (1 sheet, both sides).
Formula quality is true LaTeX: stacked fractions, real super/subscripts,
\sum, \int, \sqrt, \hat, etc.

There is a second, separate deliverable: **study-notes.pdf** (复习材料), a
readable multi-page revision document compiled from `study-notes.tex` by
`build_study_notes.py` — see the "Study notes" section at the end.

## Workflow

1. Copy `templates/cheatsheet.tex` to the session folder as `cheatsheet.tex`.
2. Replace the body (SIDE 1 and SIDE 2 sections) with the real exam content.
   **Keep the preamble unchanged.**
3. Build:

```powershell
& "<python>" "<skill-dir>/scripts/build_cheatsheet.py" `
  --tex "exam-cram/phys101/cheatsheet.tex" `
  --out "exam-cram/phys101/cheatsheet.pdf" `
  --state "exam-cram/phys101/state.json" `
  --json "exam-cram/phys101/build-report.json"
```

The build script:
- locates `tectonic` (auto-detect; or `--tectonic PATH`, or
  `$env:EXAM_CRAM_TECTONIC`);
- compiles across font sizes and keeps the **smallest font size that still
  yields exactly 2 pages** (maximises information density — the whole point
  of a cram sheet);
- renders PNG previews (needs the DSH LibreOffice Kit);
- checks every state topic id appears in the `.tex`.

## Authoring rules (the model must follow)

**The cheat sheet must be plain, uniform and dense.** The approved style:
no decoration whatsoever.

- **No decoration**: no header band, no footer, no page numbers, no boxes,
  no grey labels, no rules under titles, no bullets, no stars, no `\chk`,
  no coloured pills, no circled numbers. All of these are forbidden.
- **Uniform type**: ONE font size for everything. A topic title is a bold
  line in the same size as the body (`\topic{...}`); nothing is larger or
  smaller, nothing is grey or coloured.
- **Structure**: two A4 sides. SIDE 1 = formulas; SIDE 2 = tests, facts,
  discriminations. Each side is `\begin{multicols}{3} ... \end{multicols}`,
  separated by `\clearpage`. Do not add `\head` — there is no header.
- **Density**: a topic = one `\topic{T1 名称（权重）}` line followed by
  several sentences of inline-math (`$...$`) that flow and wrap naturally.
  Write ENOUGH content to fill both sides (each topic several lines). A
  half-empty side is a defect — add more real formulas/facts.
- Plain numbered items are written as `1)` `2)` `3)` — never `① ② ③`
  (those glyphs are missing in the Latin font).
- Escape `&` as `\&` in running text.
- If the build reports `sparse`, add more content — never decorative filler.
- Keep the `\AtBeginDocument{\fontsize{7.5pt}{8.8pt}\selectfont}` line — the
  build script rewrites it (and inserts the active size after
  `\begin{document}`).

## Reading the report

| Field | Meaning |
|---|---|
| `status` | `ok` (2 pages), `sparse` (2 pages but under-filled — add more content), `overflow` (doesn't fit at min size — trim), `ok_single` (1 page filled; content too small for 2) |
| `pages` | must be 2 |
| `size_pt` | chosen body font size (smallest that still fits 2 pages) |
| `page_fill` | fraction of each page height covered by ink; target ≥ 0.7 |
| `coverage.missing` | state topics not found in the text |
| `preview_images` | PNGs to inspect with `read_image` |

If `status` is `sparse`, **add more exam-grade content** (formulas with
conditions, worked patterns, traps, discrimination tables) and rebuild — do
not just accept a sparse sheet.

## LaTeX features supported

- true math: `\frac \dfrac \tfrac`, `^ _`, `\sqrt{}`, `\sum_{}^{}`,
  `\int_{}^{}`, `\prod`, `\hat \bar \tilde \vec \dot`, Greek
  (`\alpha \beta \gamma \mu \sigma \rho \phi \omega \Delta \Sigma \Gamma`),
  `\ln \log \exp \sin \cos`, `\le \ge \neq \approx \pm \times \cdot \equiv
  \propto \infty \partial \nabla \to \Rightarrow`,
- inline `$...$` and display `\[...\]`,
- CJK via xeCJK with Microsoft YaHei (adjust via `\setCJKmainfont` if needed),
- tables with `tabular` + `booktabs`,
- `\text{}`, `\mathrm{}` for upright text.

Not supported: TikZ pictures, `\begin{aligned}` is fine in amsmath but keep
it simple; no external images.

## Engine install (if missing)

tectonic is a single static binary (no TeX distribution needed):

```powershell
# Windows
New-Item -ItemType Directory -Force "$HOME\.local\bin" | Out-Null
iwr https://github.com/tectonic-typesetting/tectonic/releases/download/tectonic%400.17.0/tectonic-0.17.0-x86_64-pc-windows-msvc.zip -o $env:TEMP\tect.zip
Expand-Archive $env:TEMP\tect.zip -DestinationPath $env:TEMP\tect
Copy-Item "$env:TEMP\tect\tectonic.exe" "$HOME\.local\bin\tectonic.exe"
# macOS / Linux: brew install tectonic  (or apt install tectonic)
```

First run downloads the LaTeX format (a few tens of MB), cached under
`.tectonic-cache` next to the output; subsequent runs are offline.

## Troubleshooting

- `tectonic not found` → install it (above) or pass `--tectonic`.
- `no size produced a PDF` → a LaTeX error in your `.tex`. Read the printed
  engine output; most often an unbalanced `$...$`, `\frac` with wrong
  `{...}` grouping, or content outside `multicols`.
- `status: overflow` → content too long even at 5.5pt. Cut duplication,
  merge low-yield topics; keep every topic at least once.
- `status: sparse` → not enough content. Expand topics (formulas + conditions
  + traps) and rebuild.
- CJK boxes → the `xeCJK` font is set to Microsoft YaHei; on Linux/macOS
  change `\setCJKmainfont` to an installed CJK font (Noto Sans CJK SC,
  PingFang SC).

## Study notes PDF (复习材料)

Separate from the cheat sheet, the skill also produces **study-notes.pdf** —
a readable, multi-page LaTeX document built from the ingested material.

Author `study-notes.tex` from `templates/study-notes.tex`:

- `\section` per topic (or exam theme), `\subsection` for parts;
- `\defn{...}` definitions, `\align`/`\begin{align}` display formulas,
  `\begin{example}...\end{example}` worked examples, `\warn{...}` common
  errors, `\tip{...}` memory aids;
- source tags everywhere: `(L1 p.4)`, `(Slide 7)`, `(HW2 Q3)`;
- end with a "公式总表" (formula summary table) and a revision checklist;
- readable 11pt, multi-page (5–20 pages typical); **no** auto-fitting — the
  document's length is natural.

Build with:

```powershell
& "<python>" "<skill-dir>/scripts/build_study_notes.py" `
  --tex "exam-cram/phys101/study-notes.tex" `
  --out "exam-cram/phys101/study-notes.pdf" `
  --json "exam-cram/phys101/study-notes-report.json"
```

The report carries `pages` and `preview_images`; visually spot-check a few
pages (headings, formulas, tables, CJK) before delivering. `tectonic` is the
same engine used for the cheat sheet (auto-detected; see "Engine install").