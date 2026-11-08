# exam-cram — 考前抱佛脚 skill for DSH

A DSH skill that turns lecture notes, slides, homework, and sample exams into a complete last-minute study loop:

1. ingest lecture/PPT/PDF materials;
2. ingest homework and sample midterms;
3. write a concise, source-tagged summary;
4. discuss it interactively (tutoring mode);
5. run a mock exam **one question at a time**, with no answers volunteered;
6. build a priority-weighted **two-sided A4 cheat sheet** (exactly 2 pages, dense 3-column LaTeX layout) plus PNG previews.

## Install

The skill is already installed at:

```text
<workspace>/.dsh/skills/exam-cram/
```

DSH discovers project skills at `<projectRoot>/.dsh/skills` and user skills at `~/.dsh/skills`. To make it available in every workspace, copy the folder:

```powershell
Copy-Item -Recurse ".dsh/skills/exam-cram" "$HOME/.dsh/skills/exam-cram"
```

Restarting DSH is not required; the skill provider watches the roots. The `exam-cram` entry should appear in the skill catalog and can be loaded with the `skill` tool.

## Requirements

- DSH's bundled Python (`load_workspace_dependencies`) with `python-docx`, `python-pptx`, and `openpyxl` — all bundled.
- **tectonic** (or any `tectonic`-compatible LaTeX) for the cheat-sheet build. Auto-detected from `PATH`/common locations; a single static binary can be dropped next to the skill. See `references/cheatsheet-format.md` for install. TeX Live/MiKTeX also work if `tectonic` is on `PATH`.
- LibreOffice Kit for PDF page-count/preview rendering. It ships with DSH; the builder auto-detects it.
- `pdftotext` for PDF text extraction. Auto-detected from `PATH`, Git for Windows (`ucrt64`/`mingw64`/`usr`), Homebrew, and `/usr/bin`. If absent, PDFs fall back to `office-render` + vision.
- Optional: `pypdf` if already installed; not required.

## Quick start

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" init `
  --dir "exam-cram/phys101" --course "PHYS 101 Mechanics" --exam "Final" --date "2026-12-20"

& "<python>" "<skill-dir>/scripts/exam_cram.py" ingest `
  --dir "exam-cram/phys101" --source "materials"

# add topics (see templates/topics.json)
& "<python>" "<skill-dir>/scripts/exam_cram.py" add-topics `
  --dir "exam-cram/phys101" --json "topics.json"

# after the mock
& "<python>" "<skill-dir>/scripts/exam_cram.py" mastery `
  --dir "exam-cram/phys101" --write

# build the cheat sheet (tectonic LaTeX -> PDF, auto-fit to 2 pages)
& "<python>" "<skill-dir>/scripts/build_cheatsheet.py" `
  --tex "exam-cram/phys101/cheatsheet.tex" `
  --out "exam-cram/phys101/cheatsheet.pdf" `
  --state "exam-cram/phys101/state.json" `
  --json "exam-cram/phys101/build-report.json"
```

Load the skill in DSH and ask, for example:

- "用这些 lecture notes 和 sample midterm 帮我抱佛脚，最后做一张 A4 双面 cheat sheet。"
- "Quiz me one question at a time; don't show answers until I ask."
- "Rebuild the cheat sheet with more space on rotational dynamics."

## What is in the box

```text
exam-cram/
  SKILL.md                       workflow contract the agent follows
  references/
    cheatsheet-format.md         LaTeX authoring rules + build report + engine
    mock-exam.md                 blueprint, question anatomy, protocol
    ingestion.md                 source map + file-type fallbacks
    mastery.md                   mastery formula + space allocation
  templates/
    summary.md                   summary scaffold
    cheatsheet.tex               two-sided LaTeX starter (dense 3-col)
    topics.json                  topic schema
  scripts/
    exam_cram.py                 init / ingest / record / mastery / status / coverage
    build_cheatsheet.py          tectonic LaTeX -> PDF, auto-fit to 2 pages, previews
```

## Design notes

- **One question at a time** is enforced by the skill contract, not by a UI. The agent is told exactly what to show, what to record, and what never to reveal. The answer key lives in `mock/key.json`; the model may read it only when the user asks for `answer`.
- **Mastery is evidence, not a grade.** It combines weighted accuracy, coverage, and self-confidence; untested is `None`, not zero. The `space_pct` output is a budget for the cheat sheet, not a hard layout.
- **The cheat sheet is real LaTeX.** `build_cheatsheet.py` compiles `cheatsheet.tex` with tectonic, so formulas get true stacked fractions, `\sum`/`\int`/`\sqrt`, real super/subscripts, and CJK via xeCJK. The builder picks the largest font size that still fits exactly two A4 pages and reports `sparse` when the page is under-filled — density is the design goal.
- **Coverage is checked literally.** Every `state.json` topic must appear by ID or full name in `cheatsheet.tex`; the builder reports missing topics before delivery.
- **Vision is the fallback, not the default.** Text extraction is cheap and keeps page citations; rendered images are used for scanned PDFs, diagrams, and handwriting.

## Limitations

- The cheat-sheet template is an `article` + `multicol` document; exotic packages are not preloaded, but any standard AMS math works.
- The 3-column layout is balanced by multicol; very long single paragraphs may push content unevenly — the agent should inspect previews.
- Page fits are checked by page count + pixel fill and visual QA, not by a layout solver. The build script reports `sparse`/`overflow` and the agent must fix content accordingly.
- `pdftotext` extraction quality depends on the PDF's text layer. Scanned PDFs require vision and are slower.
- The mastery model is a study heuristic, not a psychometric instrument. One hard question can dominate a topic; use it to allocate space, not to make claims about ability.
