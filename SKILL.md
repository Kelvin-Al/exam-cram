---
name: exam-cram
description: "Exam cram / 考前抱佛脚 workflow: ingest lecture notes and past papers, write a concise summary, tutor interactively, run a one-question-at-a-time mock exam without revealing answers, then build a priority-weighted two-sided A4 cheat sheet. Use for finals, midterms, 期末突击, 复习提纲, or cheat-sheet requests."
whenToUse: "The user wants to cram for an exam from lecture notes/PPTs/PDFs and homework or sample exams, wants a summary, wants to be quizzed one question at a time, or wants a one-page double-sided A4 cheat sheet."
---

# Exam Cram — 考前抱佛脚

Turn a pile of lecture notes and past papers into a tight loop:

**source map → concise summary → interactive Q&A → one-question-at-a-time mock → priority-weighted 2-sided A4 cheat sheet.**

The learner should spend most of the time recalling and practising, not reading. Optimise for exam marks, not encyclopaedic coverage.

## Operating rules

1. **Work in a session folder.** Default `exam-cram/<course-slug>/` in the current workspace. Keep `state.json`, `sources.json`, `summary.md`, `mock/`, `cheatsheet.tex`, and the final PDF there. Never write into the skill directory.
2. **Never reveal a mock answer, rubric, or hint unless the user asks.** After the user answers, acknowledge only; do not say whether it is right. The next action is theirs: `answer`, `next`, `hint`, or a self-grade.
3. **Ground every claim in a source.** Tag summary claims with `[L1 p.4]`, `[Slide 7]`, `[HW2 Q3]`, or `[Sample Midterm Q1]`. Mark inferences as `(inferred)` and gaps as `(not in sources — verify)`.
4. **Exactly 2 A4 sides** for the cheat sheet. 1 page double-sided = 2 printed sides. If the builder reports 3 pages, trim content; do not deliver an overlong sheet.
5. **Cover every topic, weight the space.** A topic may be compressed to one line, but it must appear. Weak × high-weight topics get the most space.
6. **Do not interrogate.** If materials are already provided, start. Ask at most one short batch of questions when a missing fact blocks progress (exam date, allowed aids, language, question count).
7. **Reply in the user's language.** Keep course-specific technical terms in the language the course uses.

The `skill` tool reports this skill's directory. Call it `<skill-dir>` below. The Python interpreter is the `python` path returned by `load_workspace_dependencies`; call it `<python>`.

## Stage 0 — Start or resume

If the user has not given materials, ask for:

- course + exam type/date (or "unknown");
- lecture notes/PPT/PDF folder;
- homework, problem sets, sample midterms (strongly recommended);
- allowed aids (closed book? one A4? calculator? formula sheet?);
- language and time available.

Then create the session:

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" init `
  --dir "exam-cram/phys101" --course "PHYS 101 Mechanics" --exam "Final" --date "2026-12-20"
```

Resume with:

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" status --dir "exam-cram/phys101"
```

If the user only wants a cheat sheet ("just make me the sheet"), offer the fast path: ingest → summary → cheat sheet, with a short diagnostic quiz only if they want the sheet weighted to their weaknesses.

## Stage 1 — Ingest and build the exam blueprint

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" ingest `
  --dir "exam-cram/phys101" `
  --source "materials/lectures" "materials/past-papers" "materials/syllabus.pdf"
```

The script writes `sources.json` and `extracted/*.txt` with page/slide markers. It uses `python-pptx`, `python-docx`, and `pdftotext` (auto-detected; Git for Windows includes it). Then:

1. Read `sources.json`; for each source note pages, chars, and status.
2. Read the extracted text **from files**, not by dumping it into chat.
3. Handle fallbacks:
   - `scanned` / `low_text` / `no_pdf_tool` PDF: load `office-render`, render the pages, and read the images with `read_image`. For a long PDF, inspect the contents page, section openers, diagrams, and pages with formulas first.
   - `.pptx` / `.docx` with unusual content: load `office-pptx` / `office-docx` and inspect the original if extraction is incomplete.
   - Handwritten notes or photographed slides: `read_image` directly.
4. Write `source-map.md`: a table of source → pages → topics → likely exam weight, plus a "gaps" section for missing weeks/solutions.
5. Write `exam-blueprint.md`: question types, marks per type, time limit, allowed aids, recurring archetypes from homework/sample midterm, and a topic-weight table.

Ask the user to confirm the topic list and weights only if the syllabus/past papers disagree; otherwise state your assumptions and proceed. Add the topics to the session:

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" add-topics `
  --dir "exam-cram/phys101" --json "exam-cram/phys101/topics.json"
```

`topics.json` is a list of `{id, name, weight, importance, source_refs}`. Use stable IDs (`T1`, `T2`, …); `weight` can be marks, hours, or a percentage (the script normalises). `importance` is 1–3.

Set the stage:

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" stage --dir "exam-cram/phys101" --set ingest
```

## Stage 2 — Write the concise summary + study-notes PDF

Create `summary.md` from `templates/summary.md` (or the structure below). Target **1,200–2,500 words** or roughly 10–15% of the source; hard cap 3,000 words. A summary is not a transcript.

Required structure:

1. **Exam intelligence** — format, time, marks, allowed aids, question types, recurring archetypes.
2. **Coverage map** — table: topic ID, topic, weight, source refs, must-know?; every state topic appears.
3. **Core sections** — `S1`, `S1.1`, …; each has:
   - one-line idea;
   - definitions/symbols;
   - formulas with conditions;
   - one worked archetype (question → method → answer shape);
   - traps and sign/convention errors;
   - a `→ Check:` active-recall question.
4. **Formula bank** — table: formula, when to use, conditions, common mistake.
5. **Question archetypes** — table: pattern → recognition cue → standard method → marks.
6. **Gaps** — unresolved questions, conflicting sources, topics with no practice.

### 2.1 Also produce the study-notes PDF (复习材料)

The user expects a **readable, LaTeX-typeset revision document** generated from the ingested material, in addition to the summary. From the same content:

1. Copy `templates/study-notes.tex` to the session folder as `study-notes.tex` and fill it in:
   - one `\section` per topic (or exam theme), with `\defn{}` definitions, `\align` display formulas, `\begin{example}` worked examples, `\warn{}` common errors, `\tip{}` memory aids;
   - keep every source tag like `(L1 p.4)` / `(Slide 7)` / `(HW2 Q3)`;
   - a final "公式总表" (formula summary table) and a closing revision checklist;
   - readable 11pt type, multi-page (typically 5–20 pages); **no** page-fitting is applied.
2. Build it:

```powershell
& "<python>" "<skill-dir>/scripts/build_study_notes.py" `
  --tex "exam-cram/phys101/study-notes.tex" `
  --out "exam-cram/phys101/study-notes.pdf" `
  --json "exam-cram/phys101/study-notes-report.json"
```

3. The report shows `pages` and preview images; inspect a few with `read_image` to confirm headings, formulas, tables, and CJK render correctly.

Deliver `study-notes.pdf` alongside `summary.md`. This is the revision document the user reads while studying; the cheat sheet (Stage 5) is the separate last-minute sheet.

After writing, send a short message with the table of contents, the 3 highest-yield ideas, and one question: "Which section should we stress-test first?" Do **not** paste the whole summary into chat unless asked. Set stage `summary`.

## Stage 3 — Interactive tutoring

The user can ask about anything in the summary or sources. In this stage:

- Answer with `Sx.y` references and a concrete example.
- Prefer a short explanation + one check question over a lecture.
- When the user is confused, update `gaps.md` with the exact misconception, and add a note to `state.json` only if it changes a topic's status.
- If the user asks for a deeper derivation or a different explanation, give it; then return to exam framing ("what would this look like in a 6-mark question?").
- Do not start the mock exam until the user says they are ready (e.g. "quiz me", "开始模拟", "test me") or explicitly asks.

Set stage `tutor` when the summary is accepted.

## Stage 4 — Mock exam (one question at a time)

Before Q1, confirm the blueprint in **one** short message: number of questions (default 10), time (default 60 min), difficulty mix, open/closed book, calculator. Then write:

- `mock/blueprint.md` — topic × question-type × marks table, aligned to the real exam.
- `mock/questions.json` — full questions, options, marks, model answer, rubric, common errors.
- `mock/key.json` — answer key only. **Never quote this file in chat**; open it only when the user sends `answer`.
- `mock/attempts.json` — optional; `state.json` is the source of truth.

Question design: mirror the real paper's archetypes and mark distribution; sample weak topics more often, but keep high-weight topics represented. Include at least one question that combines two topics.

### Question protocol (strict)

Show **only the current question**:

```text
Q3/10 · Rotational dynamics · 8 marks · suggested 7 min

<question text, with data/options/diagram description>
```

No answer. No rubric. No "hint" unless asked. Do not show the next question in the same message.

Commands the user may send:

| User says | You do |
|---|---|
| `next` / `下一题` | Record the current question as `skip` (no score) if unanswered; show the next question only. |
| `answer` / `答案` / `解答` | Show the model answer, mark allocation, and 2–3 common errors. Do not grade the user unless asked. |
| `hint` / `提示` | Give one nudge (definition, equation to use, or first step). Never the full solution. |
| `grade 0.5` / `自评 0.5` | Record score 0.5; acknowledge briefly. |
| `pause` / `resume` | Pause; on resume restate the current question and remaining count. |
| `end` / `结束` | End the mock; show the score table and weak topics. |
| an actual answer | Say only: "Recorded. Send `answer` for the solution, `next` for the next question, or self-grade 0/0.5/1." Do **not** evaluate. |

After each command, record with:

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" record `
  --dir "exam-cram/phys101" --topic T4 --qid Q3 --type calc `
  --difficulty 4 --correct partial --score 0.5 --confidence 2 --seconds 180
```

Use `--correct skip` and omit `--score` for skipped questions. Use `--correct ungraded` if the user answered but did not want a score.

At the end, show:

- total score / attempted / skipped;
- per-topic accuracy and mastery;
- the 2–3 topics that most need cheat-sheet space;
- a one-line recommendation: "Ready for the cheat sheet?" Set stage `mock`.

## Stage 5 — Build the cheat sheet

### 5.1 Compute priorities

```powershell
& "<python>" "<skill-dir>/scripts/exam_cram.py" mastery `
  --dir "exam-cram/phys101" --write
```

This prints a table with weight, attempts, accuracy, mastery, risk, and **suggested space %**. Use it as the space budget, not as a rigid rule: a high-weight topic with no attempts is "untested", not "mastered".

### 5.2 Author `cheatsheet.tex`

Copy `templates/cheatsheet.tex` to the session folder and fill in the content. Use the rules in `references/cheatsheet-format.md`; the key points:

- **Density first.** A real cram sheet is dense: each topic gets ~10–25 short lines of formulas, conditions, worked patterns, and traps. Write flowing inline-math (`$...$`) paragraphs that naturally wrap and fill each column — do not put one formula per line.
- `\secH{3}{Title · weight}{meta}` per topic (3 = must-know red, 2 = amber, 1 = blue); `\subsec{}` for sub-blocks; `\trap` / `\tip` / `\examnt` / `\memorize` / `\noten` callouts; `\chk` for checklist items.
- **Side 1 — CORE + FORMULAS** (highest-weight must-know); **Side 2 — WEAK SPOTS + TRAPS** (low-mastery topics, traps, discrimination tables, model quick-reference, last-minute checklist).
- Every topic from `state.json` appears at least once (the builder checks this).
- Do **not** insert `\columnbreak`; let multicol balance the columns.
- Keep the `\AtBeginDocument{\fontsize{7pt}{8pt}\selectfont}` line — the builder adjusts it.
- Formulas are real LaTeX (stacked fractions, `\sum`, `\int`, `\sqrt`, `\hat`); CJK works via xeCJK.

### 5.3 Build and fit

```powershell
& "<python>" "<skill-dir>/scripts/build_cheatsheet.py" `
  --tex "exam-cram/phys101/cheatsheet.tex" `
  --out "exam-cram/phys101/cheatsheet.pdf" `
  --state "exam-cram/phys101/state.json" `
  --json "exam-cram/phys101/build-report.json"
```

The builder:

1. compiles the `.tex` with the **tectonic** LaTeX engine (auto-detected; see `references/cheatsheet-format.md` for install if missing);
2. sweeps body font size from large to small and keeps the **largest size that still gives exactly 2 A4 pages**;
3. renders PNG previews;
4. measures page fill and checks every state topic appears in the `.tex`.

Read `build-report.json`:

- `"status": "ok"`, `"pages": 2`, `page_fill ≥ 0.7` → proceed to visual check.
- `"status": "sparse"` → 2 pages but under-filled. **Add more exam-grade content** (formulas with conditions, worked patterns, traps, tables) and rebuild; do not ship a half-empty sheet.
- `"status": "overflow"` (exit 4) → too long even at the minimum size. Cut duplication, merge low-yield topics, shorten examples; keep every topic at least once. Rebuild.
- `"coverage.missing"` non-empty → add the missing topic names/IDs to `cheatsheet.tex`. Rebuild.
- `tectonic not found` (exit 2) → install tectonic (`brew install tectonic`, or download the Windows binary — link in `references/cheatsheet-format.md`), or pass `--tectonic <path>`. Do not deliver a sparse or oversized PDF.

### 5.4 Visual check (required)

Read both preview images with `read_image`:

- exactly 2 pages, both sides used;
- no clipped text at column edges or page bottom;
- no orphan section heading at the bottom of a column;
- all three columns reasonably balanced; no empty column;
- smallest text is legible when printed (≈6 pt); Chinese glyphs are not tofu;
- formulas, superscripts, subscripts, arrows, and units read correctly;
- the highest-priority topics occupy the most visual space.

If a page is underfull, add a discrimination table, a worked pattern, or a "common errors" box — not filler. If a page is overcrowded, cut repeated wording and low-yield detail, not whole topics.

### 5.5 Deliver

Keep `cheatsheet.pdf` (primary) and `cheatsheet-preview/page-*.png` (QA), and the `cheatsheet.tex` source for tweaks. Present the PDF. In the final message, state:

- exactly 2 A4 sides;
- topic coverage count (e.g. 12/12);
- the 3 topics with the most space and why;
- one line on what to do 30 minutes before the exam.

## Command reference

| Command | Purpose |
|---|---|
| `exam_cram.py init` | Create session + `state.json` |
| `exam_cram.py add-topics` | Merge topics from JSON |
| `exam_cram.py ingest` | Extract PPTX/DOCX/PDF/text into `extracted/` |
| `exam_cram.py stage` | Set `intake/ingest/summary/tutor/mock/cheatsheet/done` |
| `exam_cram.py record` | Record one mock attempt |
| `exam_cram.py mastery` | Print mastery + suggested cheat-sheet space |
| `exam_cram.py status` | Show session state and next action |
| `exam_cram.py coverage` | Check every topic appears in `cheatsheet.tex` |
| `build_cheatsheet.py` | tectonic LaTeX → PDF, auto-fit to 2 pages, render previews |
| `build_study_notes.py` | tectonic LaTeX → multi-page study-notes PDF (复习材料) |

## Quality bar

Before calling the task done:

- [ ] Every source is accounted for in `sources.json`; scanned PDFs have a vision fallback.
- [ ] Summary has section IDs, source tags, formulas with conditions, and active-recall checks.
- [ ] At least one full mock cycle happened (or the user explicitly skipped it).
- [ ] Every mock question was shown alone; no answer was volunteered.
- [ ] Every topic appears in the cheat sheet; high-weight/weak topics get more space.
- [ ] PDF is exactly 2 A4 sides and visually checked.
- [ ] Final message states coverage, page count, and the top 3 priorities.

## Files in this skill

- `references/cheatsheet-format.md` — LaTeX authoring rules, build report, engine install, troubleshooting.
- `references/mock-exam.md` — blueprint design, question writing, grading rubric, protocol examples.
- `references/ingestion.md` — file-type fallbacks and source-map template.
- `references/mastery.md` — mastery formula and how to interpret it.
- `templates/summary.md`, `templates/cheatsheet.tex`, `templates/study-notes.tex`, `templates/topics.json` — starters.
- `scripts/exam_cram.py`, `scripts/build_cheatsheet.py`, `scripts/build_study_notes.py` — deterministic helpers.
