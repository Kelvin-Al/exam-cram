# Mock exam playbook

The mock is a **recall and diagnosis tool**, not a second lecture. Its value comes from forcing retrieval, then giving feedback only when the learner asks for it.

## 1. Build the blueprint before writing questions

Read `exam-blueprint.md` and the sample midterm. Fill this table:

| Topic | Weight | Real exam archetypes | Mock Qs | Marks |
|---|---:|---|---:|---:|
| T1 Kinematics | 18% | 1D/2D constant acceleration; graph reading | 2 | 12 |
| T2 Newton's laws | 22% | FBD + incline + friction | 2 | 16 |
| … | | | | |

Rules:

- Total marks should match the real paper (or the requested practice size).
- Every topic appears at least once; high-weight and low-mastery topics get extra questions.
- Mirror question **types** (MCQ, short answer, calculation, proof, diagram, data analysis), not just topics.
- Include at least one cross-topic question and one "explain why" question.
- Difficulty mix for a 10-question set: ~3 easy, ~5 medium, ~2 hard. Do not make every question hard; the point is diagnosis.
- If the user asks for a timed paper, set the time from the real paper's marks/minute ratio, not from your own reading speed.

Save the blueprint to `mock/blueprint.md`.

## 2. Question anatomy

Each entry in `mock/questions.json`:

```json
{
  "id": "Q3",
  "topic_id": "T4",
  "type": "calc",
  "difficulty": 4,
  "marks": 8,
  "suggested_minutes": 7,
  "prompt": "A uniform solid cylinder ...",
  "options": [],
  "answer": "Full model answer with the final result boxed.",
  "rubric": [
    {"mark": 2, "criterion": "correct moment of inertia about the pivot"},
    {"mark": 3, "criterion": "torque equation and sign convention"},
    {"mark": 3, "criterion": "correct algebra and units"}
  ],
  "common_errors": [
    "using I_cm instead of the parallel-axis value",
    "forgetting the factor 1/2 in K_rot"
  ],
  "source_refs": ["L5 p.12", "HW4 Q2"]
}
```

The matching entry in `mock/key.json` can be just `{id, answer, rubric, common_errors}`. Never paste `key.json` into chat. Open it only when the user sends `answer`.

## 3. The question protocol

### First question

```text
Mock exam · 10 questions · 60 min · closed book
Blueprint: T1 2Q, T2 2Q, T3 2Q, T4 3Q, T5 1Q (weak topics weighted up)

Q1/10 · Kinematics · 6 marks · suggested 5 min

A ball is thrown vertically upward at 12 m/s. ...
```

Then stop. No answer, no rubric, no "would you like a hint?", no next question.

### After the user answers

Bad:

> Correct! The maximum height is 7.3 m. Ready for Q2?

Good:

> Recorded. Send `answer` for the solution, `next` for the next question, or self-grade 0/0.5/1.

Do not evaluate, do not confirm, do not say "close". The learner may want to revise their answer or ask for a hint.

### `answer`

```text
Q1 model answer
1. Sign convention: up +, a = −9.8 m/s².
2. v = 0 at the top → t = 12/9.8 = 1.22 s.
3. h = 12(1.22) − ½(9.8)(1.22)² = 7.35 m.

Rubric (6 marks)
- 1 mark: correct sign convention / knowns
- 2 marks: v = u + at set to zero
- 2 marks: substitution into h = ut + ½at²
- 1 mark: units and 2 s.f.

Common errors
- using g = +9.8 and getting a negative height
- forgetting the return trip when asked for total time
```

Then invite a self-grade: "Self-grade 0/0.5/1, or send `next`." If the user gives a grade, record it. If they send `next` without grading, record `ungraded` (score omitted).

### `hint`

Give one layer only:

- definition or sign convention;
- the equation family to use;
- the first step (e.g. "draw the FBD and resolve along the incline").

Never include the final answer or the full substitution. If the user asks for a second hint, give the next layer but note that hints reduce the diagnostic value.

### `next` without an answer

Record `correct: "skip"`, no score. Do not shame the learner; say "Skipped — we'll see it again in the cheat sheet." Move to the next question only.

### `pause` / `resume`

Pause the timer; on resume restate the current question number, text, and remaining time. Do not reveal anything.

### `end`

Show:

```text
Mock complete · 7/10 attempted · 42/70 marks

| Topic | Qs | Marks | Accuracy | Mastery |
|---|---:|---:|---:|---:|
| T1 Kinematics | 2 | 10/12 | 83% | 78% |
...

Weakest: T4 (rotational dynamics), T5 (oscillations)
Suggested cheat-sheet space: T4 24%, T5 21%, T2 18%
```

Then ask whether to build the cheat sheet. Set stage `mock` → `cheatsheet`.

## 4. Grading discipline

- Grade against the rubric, not against the model answer's wording.
- Award method marks for a correct setup even if arithmetic is wrong; say which mark was lost.
- If the user's answer is ambiguous, ask one clarifying question before grading.
- Do not inflate scores. A generous mock produces a misleading cheat sheet.
- When the user self-grades, accept it but record `confidence` separately (1–5). If self-grade and confidence disagree strongly, note it in `gaps.md`.

## 5. Recording

```powershell
# answered and self-graded
& "<python>" "<skill-dir>/scripts/exam_cram.py" record `
  --dir "exam-cram/phys101" --topic T1 --qid Q1 --type calc `
  --difficulty 3 --correct yes --score 1 --confidence 4 --seconds 95

# answered but not graded
... --correct ungraded --confidence 2

# skipped
... --correct skip
```

`--difficulty` is 1–5 and weights the accuracy calculation; use the question's cognitive demand, not how hard it felt to the learner.

## 6. Question-quality checks

Before showing Q1:

- [ ] The prompt is self-contained: all data, units, and assumptions are present.
- [ ] The marks and suggested time are stated.
- [ ] The answer is unambiguous and the rubric sums to the marks.
- [ ] The question is answerable from the provided sources, not from outside knowledge.
- [ ] No answer, rubric, or "the answer is" text leaks into the prompt.
- [ ] The next question is not visible in the same message.
- [ ] `questions.json` and `key.json` agree on IDs and answers.

## 7. Adapting when the user is struggling

- Two consecutive failures on the same topic: stop the mock, give a 3-minute micro-lesson, then ask one similar question. Record both attempts separately.
- The user asks "just show me the answers": switch to review mode, show answers one by one, and record `ungraded`; do not pretend this is a mock.
- The user is short on time: run a 5-question "high-yield only" mock weighted to the top 3 topics, then build the sheet.
- The user wants a full timed paper: generate all questions first, hide them, and reveal one at a time. Never dump the paper.
