# Ingestion playbook

## Goal

Build a **source map** and an **exam blueprint** before writing a single summary sentence. The summary is only as good as the topic/weight model behind it.

## 1. Inventory

Put lecture notes, slides, handouts, homework, sample exams, and the syllabus in one folder tree:

```text
materials/
  lectures/       lecture01.pdf, lecture02.pptx, ...
  homework/       hw01.pdf, hw02.docx, ...
  past/           sample-midterm.pdf, final-2024.pdf, ...
  syllabus.pdf
```

Run:

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" ingest `
  --dir "exam-cram/phys101" --source "materials"
```

The script writes:

- `sources.json` — one record per file: path, kind, pages, chars, text file, status.
- `extracted/<name>.txt` — text with `## Page N` / `## Slide N` markers and a source header.

Read the inventory, not the whole corpus in chat. Spot-check one extracted file per source type.

## 2. Status handling

| Status | Meaning | Next step |
|---|---|---|
| `ok` | Usable text | Read the extracted file; use page markers for citations. |
| `low_text` | Some text, likely sparse/short | Read it; if important, render pages for diagrams. |
| `scanned` | PDF has no text layer | `office-render` pages → `read_image`; OCR only if necessary. |
| `no_pdf_tool` | `pdftotext` missing | Install nothing; use `office-render` + vision, or `pypdf` if already present. |
| `error` | Parse failed | Open the original with the matching office skill; report the failure if it blocks. |
| `image` | Handwritten/photo | `read_image`; ask the user to clarify illegible parts. |

### PDF fallback

`pdftotext` is auto-detected from `PATH`, Git for Windows (`ucrt64`, `mingw64`, `usr`), Homebrew, and `/usr/bin`. If it is unavailable:

1. Load `office-render`.
2. Render the relevant pages at the default DPI (or 144 DPI for dense text).
3. Read the PNGs with `read_image`.
4. For a 40+ page deck, inspect the contents page, section openers, and pages with formulas/diagrams first; do not read every page blindly.

### PPTX/DOCX fallback

`python-pptx`/`python-docx` extract text, tables, notes, and a marker for images/charts, but not SmartArt, equations as LaTeX, or embedded objects. If a slide is mostly a diagram:

- load `office-pptx` / `office-docx` and inspect the shape tree;
- or render the slide/page with `office-render` and read the image;
- describe the diagram in `source-map.md` and use it in the summary.

### Equations

DOCX equations are extracted as `[math: ...]` text from OMML runs. It is linear and lossy. For formula-heavy courses:

- check the equations against the rendered page;
- rewrite them in the summary/cheat sheet in standard notation;
- do not trust an extracted equation blindly.

## 3. Source map

Write `source-map.md`:

```markdown
# Source map — PHYS 101

| ID | Source | Pages/slides | Topics | Likely weight | Quality |
|---|---|---|---:|---:|---|
| L1 | lecture01.pdf | 1–18 | kinematics, projectiles | high | text |
| L2 | lecture02.pptx | 1–24 | Newton's laws, friction | high | text + diagrams |
| HW3 | homework03.docx | — | rotation | medium | text |
| SM | sample-midterm.pdf | 1–6 | all | high | text |

## Gaps
- No lecture on oscillations after L8; only the textbook chapter.
- Homework 5 has no solutions; weights inferred from marks.
- Diagrams on L2 slides 9–11 need visual inspection.
```

## 4. Exam blueprint

Write `exam-blueprint.md`:

```markdown
# Exam blueprint — PHYS 101 Final

- Date: 2026-12-20, 09:00, 120 min
- Format: 6 short answers (30%), 4 calculations (50%), 1 design/explain (20%)
- Allowed: one A4 double-sided formula sheet, non-programmable calculator
- Recurring archetypes:
  1. FBD + incline + friction (almost every year)
  2. Work–energy vs Newton's laws choice
  3. Collision + energy loss
  4. Rotation about a pivot (parallel-axis trap)
- Topic weights (from syllabus hours × past-paper marks):
  | Topic | Weight | Evidence |
  |---|---:|---|
  | Kinematics | 18% | L1–L2, SM Q1 |
  ...
```

If the syllabus and past papers disagree, use the past papers for question style and the syllabus for coverage; state the assumption.

## 5. Topic list

Create `topics.json`:

```json
[
  {"id": "T1", "name": "Kinematics", "weight": 18, "importance": 3,
   "source_refs": ["L1 p.1-18", "SM Q1"]},
  {"id": "T2", "name": "Newton's laws", "weight": 22, "importance": 3,
   "source_refs": ["L2", "HW2"]}
]
```

- `weight`: marks, hours, or percentage; normalised later. Use the best evidence you have and note it.
- `importance`: 1–3; 3 = almost certain to appear, 2 = likely, 1 = peripheral.
- `source_refs`: short, stable tags used in the summary.

## 6. What to do with a huge corpus

- Prioritise: syllabus > past papers > homework > lecture slides > textbook.
- Cluster by topic, not by file; duplicate coverage is fine, but do not let one 80-slide deck dominate.
- Read the first 2 and last 2 slides of each deck; the middle is often examples.
- For each topic, find at least one **worked example** and one **exam-style question**; those become the summary's archetype and the mock's template.
- Stop ingesting when new sources stop changing the topic/weight table. Note what you skipped in `source-map.md`.
