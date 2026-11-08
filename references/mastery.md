# Mastery model and space allocation

The point of this model is not psychometric accuracy. It is to answer one question: **which topics deserve the scarce space on a two-sided A4 sheet?**

## Inputs

- `weight`: topic weight from marks/hours/syllabus; normalised to sum 1.
- `importance`: 1–3; raises the target number of attempts.
- attempts: `score` (0–1 or null), `difficulty` (1–5), `confidence` (0–5), `ts`.
- `correct`: `yes` / `no` / `partial` / `skip` / `ungraded` (for display; `score` drives the math).

## Formula

For each topic:

```
graded = attempts with a numeric score
accuracy = Σ(score × difficulty × recency) / Σ(difficulty × recency)
coverage = min(1, len(graded) / target),  target = 3 if importance = 3 else 2
confidence = mean(confidence) / 5

if graded:
    mastery = 0.70 × accuracy + 0.15 × confidence + 0.15 × coverage
elif attempts:
    mastery = 0.20 + 0.10 × coverage      # answered but ungraded
else:
    mastery = None                        # untested
```

- `recency` = 1.0 for the last 5 graded attempts, 0.7 for older ones.
- `difficulty` weights harder questions more; a correct hard question is stronger evidence than a correct easy one.
- `confidence` falls back to `accuracy` when the learner did not self-rate.
- `mastery = None` is **not** zero. It means "no evidence". It is displayed as `—` and treated as 0.20 for risk, so untested topics surface for diagnosis.

## Risk and space

```
risk      = weight × (1 − effective_mastery)
priority  = 0.55 × weight + 0.45 × (1 − effective_mastery)
space_pct = priority normalised to 100, floor 4–6%, cap 24%
```

`effective_mastery` is `mastery` or 0.20 for untested topics.

Interpretation:

- **High weight + low mastery** → biggest space allocation; deserves a full section with formulas, a worked pattern, and traps.
- **High weight + high mastery** → keep a compact formula/definition block; do not delete it.
- **Low weight + low mastery** → a short "don't forget" line or a table row; do not spend a column.
- **Untested** → treat as uncertain, not weak. Give it enough space to be useful, and flag it as untested in the final message.

Space percentages are a **budget**, not a layout instruction. The builder fills columns sequentially; the model decides how many lines each section gets.

## Status labels

| Mastery | Status | Meaning |
|---:|---|---|
| `None` | `untested` | No graded attempt |
| < 0.45 | `weak` | Likely to lose marks; needs a worked pattern |
| 0.45–0.65 | `shaky` | Knows the idea, loses execution marks |
| ≥ 0.65 | `solid` | Maintain with a compact reminder |

## Worked example

| Topic | Weight | Attempts | Accuracy | Mastery | Risk | Space |
|---|---:|---:|---:|---:|---:|---:|
| Rotation | 20% | 1 fail (hard) | 0% | 8% | 18% | 24% |
| Newton | 22% | 0 | — | — | 18% | 22% |
| Kinematics | 18% | 2 (1 right, 1 half) | 71% | 70% | 5% | 12% |

The sheet gives rotation a full column, Newton a full section with a diagnostic question, and kinematics a compact formula block. That is the intended behaviour: **space follows risk, not chapter order.**

## Commands

```powershell
# record one attempt
& "<python>" "<skill-dir>/scripts/exam_cram.py" record `
  --dir "exam-cram/phys101" --topic T4 --qid Q3 --type calc `
  --difficulty 5 --correct no --confidence 1 --seconds 200

# compute and save mastery.json
& "<python>" "<skill-dir>/scripts/exam_cram.py" mastery `
  --dir "exam-cram/phys101" --write
```

`mastery.json` contains one row per topic with `weight`, `attempts`, `graded`, `accuracy`, `confidence`, `coverage`, `mastery`, `risk`, `priority`, `status`, and `space_pct`.

## Caveats

- One hard question can dominate a topic's mastery. If a topic has only one attempt, read it as a signal, not a verdict.
- Self-confidence is often miscalibrated. If confidence is high but accuracy is low, note "overconfident" in `gaps.md`; the cheat sheet should include a discrimination/trap box.
- Skipped questions are not failures; they are missing evidence. Record them as `skip` and let the mock blueprint sample the topic again.
- Do not use mastery to justify dropping a topic from the sheet. Coverage is mandatory; only space is variable.
