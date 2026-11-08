# exam-cram — 考前抱佛脚 skill for DSH

A DSH skill that turns lecture notes, slides, homework, and sample exams into a complete last-minute study loop:

1. ingest lecture/PPT/PDF materials;
2. ingest homework and sample midterms;
3. write a concise, source-tagged summary;
4. discuss it interactively (tutoring mode);
5. run a mock exam **one question at a time**, with no answers volunteered;
6. build a priority-weighted **two-sided A4 cheat sheet** (exactly 2 pages), plus an editable DOCX and PNG previews.

## Install

> This is a [DeepSeek Harness (DSH)](https://www.deepseek.com/) skill, not a standalone application. Install DSH first.

### From GitHub (recommended)

Clone this repository into DSH's user skills directory so the skill is available in every workspace:

```powershell
# Windows (PowerShell)
New-Item -ItemType Directory -Force "$HOME\.dsh\skills" | Out-Null
git clone https://github.com/<you>/exam-cram.git "$HOME\.dsh\skills\exam-cram"
```

```bash
# macOS / Linux
mkdir -p ~/.dsh/skills
git clone https://github.com/<you>/exam-cram.git ~/.dsh/skills/exam-cram
```

### From a local copy / zip

DSH discovers project skills at `<projectRoot>/.dsh/skills` and user skills at `~/.dsh/skills`. Copy the folder (or unzip) so that `SKILL.md` sits directly at:

```text
<workspace>/.dsh/skills/exam-cram/SKILL.md      # project-scoped
~/.dsh/skills/exam-cram/SKILL.md                # user-scoped (all workspaces)
```

```powershell
Copy-Item -Recurse ".dsh/skills/exam-cram" "$HOME/.dsh/skills/exam-cram"
```

Restarting DSH is not required; the skill provider watches the roots. Open a new session and the `exam-cram` entry should appear in the skill catalog; load it with the `skill` tool or just describe a cram request.

## Requirements

- DSH's bundled Python (`load_workspace_dependencies`) with `python-docx`, `python-pptx`, and `openpyxl` — all bundled.
- LibreOffice Kit for DOCX → PDF conversion and page rendering. It ships with DSH; the builder auto-detects it. If it is missing, the builder emits a DOCX and the `office-export-pdf` / `office-render` skills can finish the job.
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

# build the cheat sheet
& "<python>" "<skill-dir>/scripts/build_cheatsheet.py" `
  --md "exam-cram/phys101/cheatsheet.md" `
  --out "exam-cram/phys101/cheatsheet.pdf" `
  --docx "exam-cram/phys101/cheatsheet.docx" `
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
    cheatsheet-format.md         Markdown dialect + layout + QA
    mock-exam.md                 blueprint, question anatomy, protocol
    ingestion.md                 source map + file-type fallbacks
    mastery.md                   mastery formula + space allocation
  templates/
    summary.md                   summary scaffold
    cheatsheet.md                two-sided starter
    topics.json                  topic schema
  scripts/
    exam_cram.py                 init / ingest / record / mastery / status / coverage
    build_cheatsheet.py          DOCX -> PDF, fit to 2 pages, render previews
```

## Design notes

- **One question at a time** is enforced by the skill contract, not by a UI. The agent is told exactly what to show, what to record, and what never to reveal. The answer key lives in `mock/key.json`; the model may read it only when the user asks for `answer`.
- **Mastery is evidence, not a grade.** It combines weighted accuracy, coverage, and self-confidence; untested is `None`, not zero. The `space_pct` output is a budget for the cheat sheet, not a hard layout.
- **The cheat sheet is generated from Markdown.** A deterministic builder handles 3-column A4 layout, CJK fonts, formula conversion, table widths, header bands, and the 2-page fit loop. This is more reliable than asking a model to hand-place text in Word.
- **Coverage is checked literally.** Every `state.json` topic must appear by ID or full name in `cheatsheet.md`; the builder reports missing topics before delivery.
- **Vision is the fallback, not the default.** Text extraction is cheap and keeps page citations; rendered images are used for scanned PDFs, diagrams, and handwriting.

## Limitations

- The builder's math support is LaTeX-lite (common symbols, fractions, roots, super/subscripts), not a full TeX engine. Matrices, aligned environments, and chemical equations need plain-text rewrites.
- The 3-column layout works best with 10–18 topics and 4–10 lines each. 40+ dense sections will overflow; the builder reports `overflow` rather than silently producing 3 pages.
- Page fits are checked by page count and visual QA, not by a layout solver. The agent must inspect the rendered previews for clipping, orphans, and column balance.
- `pdftotext` extraction quality depends on the PDF's text layer. Scanned PDFs require vision and are slower.
- The mastery model is a study heuristic, not a psychometric instrument. One hard question can dominate a topic; use it to allocate space, not to make claims about ability.
