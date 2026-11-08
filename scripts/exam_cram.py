#!/usr/bin/env python3
"""Exam Cram session manager.

Creates a study session, extracts text from lecture/homework files, records
mock-exam attempts, and computes a priority-weighted mastery table used by the
two-sided A4 cheat-sheet builder.

Run with the Python interpreter returned by `load_workspace_dependencies`.
Only python-docx, python-pptx and openpyxl are required (all bundled with DSH).
"""
from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.oxml.ns import qn
    from docx.table import Table
    from docx.text.paragraph import Paragraph
except Exception:  # pragma: no cover - reported at runtime
    Document = None
    Table = None
    Paragraph = None
    qn = None

try:
    from pptx import Presentation
except Exception:  # pragma: no cover
    Presentation = None

try:
    import openpyxl
except Exception:  # pragma: no cover
    openpyxl = None

STAGES = ["intake", "ingest", "summary", "tutor", "mock", "cheatsheet", "done"]
SUPPORTED = {".pptx", ".docx", ".pdf", ".md", ".txt", ".markdown", ".html",
             ".htm", ".xlsx", ".png", ".jpg", ".jpeg", ".webp", ".gif"}
TEXT_EXT = {".md", ".txt", ".markdown"}


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def now() -> str:
    return _dt.datetime.now().astimezone().isoformat(timespec="seconds")


def die(msg: str, code: int = 1):
    print(f"error: {msg}", file=sys.stderr)
    raise SystemExit(code)


def read_json(path: Path, default=None):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        die(f"cannot parse {path}: {exc}")


def write_json(path: Path, obj):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def slugify(text: str, fallback: str = "session") -> str:
    text = (text or "").strip().lower()
    text = re.sub(r"[^\w\u4e00-\u9fff-]+", "-", text, flags=re.UNICODE)
    text = re.sub(r"-{2,}", "-", text).strip("-")
    return text[:60] or fallback


def norm_name(text: str) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", (text or "").lower()).strip()


def session_paths(root: Path):    return {
        "state": root / "state.json",
        "sources": root / "sources.json",
        "extracted": root / "extracted",
        "mock": root / "mock",
        "preview": root / "cheatsheet-preview",
        "summary": root / "summary.md",
        "gaps": root / "gaps.md",
        "cheatsheet_md": root / "cheatsheet.md",
        "cheatsheet_tex": root / "cheatsheet.tex",
        "cheatsheet_pdf": root / "cheatsheet.pdf",
        "cheatsheet_docx": root / "cheatsheet.docx",
    }


def load_state(root: Path):
    path = session_paths(root)["state"]
    state = read_json(path)
    if not state:
        die(f"no state.json in {root}; run `init` first")
    return state


def save_state(root: Path, state):
    state["updated_at"] = now()
    write_json(session_paths(root)["state"], state)


def normalise_weight(value) -> float:
    try:
        w = float(value)
    except (TypeError, ValueError):
        return 0.0
    return max(0.0, w)


def normalised_weights(topics):
    raw = [normalise_weight(t.get("weight", 1.0)) for t in topics]
    total = sum(raw)
    if total <= 0:
        return [1.0 / len(topics)] * len(topics) if topics else []
    return [w / total for w in raw]


def find_pdftotext():
    cands = []
    if sys.platform.startswith("win"):
        for env in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
            base = os.environ.get(env)
            if base:
                cands += [
                    Path(base) / "Git" / "ucrt64" / "bin" / "pdftotext.exe",
                    Path(base) / "Git" / "mingw64" / "bin" / "pdftotext.exe",
                    Path(base) / "Git" / "usr" / "bin" / "pdftotext.exe",
                ]
        cands += [
            Path(r"C:\Program Files\Git\ucrt64\bin\pdftotext.exe"),
            Path(r"C:\Program Files\Git\mingw64\bin\pdftotext.exe"),
            Path(r"C:\Program Files\Git\usr\bin\pdftotext.exe"),
        ]
    else:
        cands += [
            Path("/opt/homebrew/bin/pdftotext"),
            Path("/usr/local/bin/pdftotext"),
            Path("/usr/bin/pdftotext"),
        ]
    which = shutil.which("pdftotext")
    if which:
        cands.insert(0, Path(which))
    for p in cands:
        if p.exists():
            return p
    return None


def decode_text_bytes(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-8", "gb18030", "big5", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace")


def iter_docx_blocks(doc):
    """Yield paragraphs and tables in document order."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn("w:p"):
            yield Paragraph(child, doc)
        elif child.tag == qn("w:tbl"):
            yield Table(child, doc)


# --------------------------------------------------------------------------- #
# extraction
# --------------------------------------------------------------------------- #
def extract_pptx(path: Path, out: Path):
    if Presentation is None:
        return None, "python-pptx unavailable", "error"
    prs = Presentation(str(path))
    lines = [f"# Source: {path}", f"# Kind: pptx", f"# Slides: {len(prs.slides)}", ""]
    for idx, slide in enumerate(prs.slides, 1):
        lines.append(f"## Slide {idx}")
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False):
                text = shape.text_frame.text.strip()
                if text:
                    lines.append(text)
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    lines.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
            if getattr(shape, "has_chart", False):
                lines.append("[chart]")
            try:
                if shape.shape_type == 13:  # picture
                    lines.append("[image]")
            except Exception:
                pass
        try:
            notes = slide.notes_slide.notes_text_frame.text.strip()
            if notes:
                lines.append(f"[notes] {notes}")
        except Exception:
            pass
        lines.append("")
    text = "\n".join(lines)
    out.write_text(text, encoding="utf-8")
    return len(prs.slides), text, "ok"


def extract_docx(path: Path, out: Path):
    if Document is None:
        return None, "python-docx unavailable", "error"
    doc = Document(str(path))
    lines = [f"# Source: {path}", "# Kind: docx", ""]
    images = 0
    for block in iter_docx_blocks(doc):
        if isinstance(block, Paragraph):
            text = block.text.strip()
            try:
                math = "".join(t.text or "" for t in block._p.xpath(".//m:t"))
            except Exception:
                math = ""
            if math:
                text = (text + f" [math: {math}]").strip()
            if text:
                lines.append(text)
            images += len(block._p.xpath(".//w:drawing"))
        else:
            for row in block.rows:
                lines.append(" | ".join(c.text.strip().replace("\n", " ") for c in row.cells))
            lines.append("")
    if images:
        lines.append(f"[images: {images}]")
    text = "\n".join(lines)
    out.write_text(text, encoding="utf-8")
    return None, text, "ok"


def extract_pdf(path: Path, out: Path):
    tool = find_pdftotext()
    tmp = out.with_suffix(".raw.txt")
    if tool:
        for layout in (True, False):
            cmd = [str(tool), "-enc", "UTF-8"]
            if layout:
                cmd.append("-layout")
            cmd += [str(path), str(tmp)]
            try:
                res = subprocess.run(cmd, capture_output=True, text=True,
                                     encoding="utf-8", errors="replace", timeout=180)
            except Exception as exc:
                return None, f"pdftotext failed: {exc}", "error"
            if res.returncode == 0 and tmp.exists() and tmp.stat().st_size > 0:
                raw = decode_text_bytes(tmp.read_bytes())
                pages = raw.count("\f") + (1 if raw.strip() else 0)
                chunks = raw.split("\f")
                lines = [f"# Source: {path}", "# Kind: pdf", f"# Pages: {pages}", ""]
                for idx, chunk in enumerate(chunks, 1):
                    chunk = chunk.strip("\n")
                    if chunk.strip():
                        lines.append(f"## Page {idx}")
                        lines.append(chunk.rstrip())
                        lines.append("")
                text = "\n".join(lines)
                out.write_text(text, encoding="utf-8")
                tmp.unlink(missing_ok=True)
                status = "ok" if len(text) > 400 else "low_text"
                return pages, text, status
        tmp.unlink(missing_ok=True)
        return None, "pdftotext produced no text", "scanned"
    # no pdftotext: try pypdf if the environment already has it
    try:
        from pypdf import PdfReader  # type: ignore
    except Exception:
        return None, "pdftotext not found; PDF is image-only or needs office-render + vision", "no_pdf_tool"
    try:
        reader = PdfReader(str(path))
        lines = [f"# Source: {path}", "# Kind: pdf", f"# Pages: {len(reader.pages)}", ""]
        for idx, page in enumerate(reader.pages, 1):
            text = page.extract_text() or ""
            if text.strip():
                lines.append(f"## Page {idx}")
                lines.append(text.strip())
                lines.append("")
        text = "\n".join(lines)
        out.write_text(text, encoding="utf-8")
        return len(reader.pages), text, "ok" if len(text) > 400 else "low_text"
    except Exception as exc:
        return None, f"PDF extraction failed: {exc}", "error"


def extract_text_file(path: Path, out: Path):
    text = decode_text_bytes(path.read_bytes())
    out.write_text(f"# Source: {path}\n# Kind: text\n\n{text}", encoding="utf-8")
    return None, text, "ok"


def extract_html(path: Path, out: Path):
    raw = decode_text_bytes(path.read_bytes())
    try:
        from lxml import html
        tree = html.fromstring(raw)
        text = "\n".join(t.strip() for t in tree.text_content().splitlines() if t.strip())
    except Exception:
        text = re.sub(r"<[^>]+>", " ", raw)
        text = re.sub(r"\s+", " ", text)
    out.write_text(f"# Source: {path}\n# Kind: html\n\n{text}", encoding="utf-8")
    return None, text, "ok"


def extract_xlsx(path: Path, out: Path):
    if openpyxl is None:
        return None, "openpyxl unavailable", "error"
    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    lines = [f"# Source: {path}", "# Kind: xlsx", ""]
    for ws in wb.worksheets:
        lines.append(f"## Sheet: {ws.title}")
        for r_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
            if r_idx > 300:
                lines.append("[truncated after 300 rows]")
                break
            cells = ["" if v is None else str(v) for v in row]
            if any(c.strip() for c in cells):
                lines.append(" | ".join(cells))
        lines.append("")
    wb.close()
    text = "\n".join(lines)
    out.write_text(text, encoding="utf-8")
    return None, text, "ok"


def extract_one(path: Path, out_dir: Path):
    suffix = path.suffix.lower()
    stem = slugify(path.stem, "source")
    out = out_dir / f"{stem}.txt"
    n = 1
    while out.exists():
        n += 1
        out = out_dir / f"{stem}-{n}.txt"
    if suffix == ".pptx":
        pages, text, status = extract_pptx(path, out)
    elif suffix == ".docx":
        pages, text, status = extract_docx(path, out)
    elif suffix == ".pdf":
        pages, text, status = extract_pdf(path, out)
    elif suffix in TEXT_EXT:
        pages, text, status = extract_text_file(path, out)
    elif suffix in {".html", ".htm"}:
        pages, text, status = extract_html(path, out)
    elif suffix == ".xlsx":
        pages, text, status = extract_xlsx(path, out)
    elif suffix in {".png", ".jpg", ".jpeg", ".webp", ".gif"}:
        out.write_text(f"# Source: {path}\n# Kind: image\n\n[image-only source; inspect with read_image]\n",
                       encoding="utf-8")
        pages, text, status = None, "", "image"
    else:
        return None
    chars = len(text or "")
    return {
        "path": str(path),
        "kind": suffix.lstrip("."),
        "pages": pages,
        "chars": chars,
        "text_file": str(out),
        "status": status,
        "extracted_at": now(),
    }


# --------------------------------------------------------------------------- #
# commands
# --------------------------------------------------------------------------- #
def cmd_init(args):
    root = Path(args.dir).expanduser().resolve()
    sp = session_paths(root)
    if sp["state"].exists() and not args.force:
        die(f"{sp['state']} already exists; use --force to reinitialise")
    for key in ("extracted", "mock", "preview"):
        sp[key].mkdir(parents=True, exist_ok=True)
    state = {
        "version": 1,
        "course": args.course,
        "exam": args.exam,
        "exam_date": args.date or "",
        "language": args.language,
        "stage": "intake",
        "created_at": now(),
        "updated_at": now(),
        "topics": [],
        "attempts": [],
        "notes": {},
    }
    if args.topics:
        topics = read_json(Path(args.topics))
        state["topics"] = topics.get("topics", topics) if isinstance(topics, dict) else topics
    save_state(root, state)
    readme = root / "README.md"
    if not readme.exists():
        readme.write_text(
            f"# {args.course} — {args.exam}\n\n"
            f"Session created {now()}. Stage: intake.\n\n"
            "## Files\n"
            "- `state.json` — topics, attempts, stage\n"
            "- `sources.json` — extracted source inventory\n"
            "- `summary.md` — concise cram summary\n"
            "- `mock/` — blueprint, questions, answer key, attempts\n"
            "- `cheatsheet.md` — authored source for the cheat sheet\n"
            "- `cheatsheet.pdf` — final 2-sided A4 deliverable\n\n"
            "Never put the mock answer key in the chat.\n",
            encoding="utf-8",
        )
    print(json.dumps({"session": str(root), "state": str(sp["state"]),
                      "topics": len(state["topics"]), "stage": state["stage"]},
                     ensure_ascii=False, indent=2))


def cmd_add_topics(args):
    root = Path(args.dir).expanduser().resolve()
    state = load_state(root)
    payload = read_json(Path(args.json))
    topics = payload.get("topics", payload) if isinstance(payload, dict) else payload
    if not isinstance(topics, list):
        die("topics JSON must be a list or an object with a `topics` list")
    existing = {t.get("id") for t in state["topics"]}
    added = 0
    for t in topics:
        if not isinstance(t, dict) or not t.get("id") or not t.get("name"):
            die("each topic needs `id` and `name`")
        if t["id"] in existing:
            continue
        state["topics"].append({
            "id": str(t["id"]),
            "name": str(t["name"]),
            "weight": normalise_weight(t.get("weight", 1.0)),
            "importance": int(t.get("importance", 2)),
            "source_refs": t.get("source_refs", []),
            "notes": t.get("notes", ""),
        })
        existing.add(t["id"])
        added += 1
    save_state(root, state)
    print(json.dumps({"added": added, "total": len(state["topics"])}, ensure_ascii=False, indent=2))


def cmd_stage(args):
    root = Path(args.dir).expanduser().resolve()
    state = load_state(root)
    if args.set not in STAGES:
        die(f"stage must be one of: {', '.join(STAGES)}")
    state["stage"] = args.set
    save_state(root, state)
    print(json.dumps({"stage": state["stage"]}, ensure_ascii=False, indent=2))


def cmd_record(args):
    root = Path(args.dir).expanduser().resolve()
    state = load_state(root)
    topic_ids = {t["id"] for t in state["topics"]}
    if args.topic not in topic_ids:
        die(f"unknown topic {args.topic!r}; known: {', '.join(sorted(topic_ids)) or '(none)'}")
    if args.score is None:
        score = {"yes": 1.0, "no": 0.0, "partial": 0.5}.get(args.correct)
    else:
        score = max(0.0, min(1.0, float(args.score)))
    attempt = {
        "qid": args.qid or f"Q{len(state['attempts']) + 1}",
        "topic_id": args.topic,
        "type": args.type,
        "difficulty": max(1, min(5, int(args.difficulty))),
        "correct": args.correct,
        "score": score,
        "confidence": None if args.confidence is None else max(0, min(5, int(args.confidence))),
        "seconds": args.seconds,
        "notes": args.notes or "",
        "ts": now(),
    }
    state["attempts"].append(attempt)
    save_state(root, state)
    print(json.dumps({"recorded": attempt, "attempts": len(state["attempts"])},
                     ensure_ascii=False, indent=2))


def compute_mastery(state):
    topics = state.get("topics", [])
    attempts = state.get("attempts", [])
    weights = normalised_weights(topics)
    rows = []
    for topic, weight in zip(topics, weights):
        atts = [a for a in attempts if a.get("topic_id") == topic["id"]]
        graded = [a for a in atts if a.get("score") is not None]
        acc = None
        if graded:
            num = den = 0.0
            for idx, a in enumerate(graded):
                diff = max(1, int(a.get("difficulty", 3)))
                recency = 1.0 if idx >= len(graded) - 5 else 0.7
                w = diff * recency
                num += float(a.get("score", 0)) * w
                den += w
            acc = num / den if den else None
        cov = min(1.0, len(graded) / (3 if int(topic.get("importance", 2)) >= 3 else 2))
        confs = [float(a["confidence"]) / 5.0 for a in atts if a.get("confidence") is not None]
        conf = sum(confs) / len(confs) if confs else None
        if acc is not None:
            mastery = 0.70 * acc + 0.15 * (conf if conf is not None else acc) + 0.15 * cov
        elif atts:
            mastery = 0.20 + 0.10 * cov
        else:
            mastery = None
        effective = 0.20 if mastery is None else mastery
        risk = weight * (1.0 - effective)
        priority = 0.55 * weight + 0.45 * (1.0 - effective)
        if mastery is None:
            status = "untested"
        elif mastery < 0.45:
            status = "weak"
        elif mastery < 0.65:
            status = "shaky"
        else:
            status = "solid"
        rows.append({
            "id": topic["id"],
            "name": topic["name"],
            "weight": weight,
            "weight_raw": topic.get("weight", 1.0),
            "attempts": len(atts),
            "graded": len(graded),
            "accuracy": None if acc is None else round(acc, 3),
            "confidence": None if conf is None else round(conf, 3),
            "coverage": round(cov, 3),
            "mastery": None if mastery is None else round(mastery, 3),
            "risk": round(risk, 4),
            "priority": round(priority, 4),
            "status": status,
        })
    total_priority = sum(r["priority"] for r in rows) or 1.0
    min_space = 6.0 if len(rows) <= 10 else 4.0
    for r in rows:
        r["space_pct"] = round(max(min_space, 100.0 * r["priority"] / total_priority), 1)
    # normalise again after the floor and cap at 24%
    for _ in range(3):
        total = sum(r["space_pct"] for r in rows) or 1.0
        for r in rows:
            r["space_pct"] = min(24.0, r["space_pct"] * 100.0 / total)
    total = sum(r["space_pct"] for r in rows) or 1.0
    for r in rows:
        r["space_pct"] = round(r["space_pct"] * 100.0 / total, 1)
    rows.sort(key=lambda r: (-r["risk"], -r["weight"]))
    return rows


def format_mastery(rows):
    header = "| ID | Topic | Wt | Att | Acc | Mastery | Risk | Space | Status |"
    sep = "|---|---|---:|---:|---:|---:|---:|---:|---|"
    lines = [header, sep]
    for r in rows:
        acc = "—" if r["accuracy"] is None else f"{r['accuracy']*100:.0f}%"
        mastery = "—" if r["mastery"] is None else f"{r['mastery']*100:.0f}%"
        lines.append(
            f"| {r['id']} | {r['name']} | {r['weight']*100:.0f}% | {r['attempts']} | "
            f"{acc} | {mastery} | {r['risk']*100:.0f}% | {r['space_pct']:.0f}% | {r['status']} |"
        )
    return "\n".join(lines)


def cmd_mastery(args):
    root = Path(args.dir).expanduser().resolve()
    state = load_state(root)
    rows = compute_mastery(state)
    print(format_mastery(rows))
    recs = [r for r in rows if r["status"] in {"weak", "untested"}][:5]
    if recs:
        print("\nTop priorities:")
        for r in recs:
            print(f"- {r['id']} {r['name']}: {r['status']}, {r['weight']*100:.0f}% weight, "
                  f"~{r['space_pct']:.0f}% of the cheat sheet")
    if args.write:
        out = root / "mastery.json"
        write_json(out, {"generated_at": now(), "course": state.get("course"),
                         "exam": state.get("exam"), "rows": rows})
        print(f"\nwrote {out}")
    if args.json:
        print(json.dumps(rows, ensure_ascii=False, indent=2))


def cmd_status(args):
    root = Path(args.dir).expanduser().resolve()
    state = load_state(root)
    sp = session_paths(root)
    sources = read_json(sp["sources"], {"items": []}) or {"items": []}
    items = sources.get("items", [])
    rows = compute_mastery(state)
    counts = {s: sum(1 for r in rows if r["status"] == s) for s in ("weak", "shaky", "solid", "untested")}
    info = {
        "course": state.get("course"),
        "exam": state.get("exam"),
        "stage": state.get("stage"),
        "topics": len(state.get("topics", [])),
        "attempts": len(state.get("attempts", [])),
        "sources": len(items),
        "source_errors": [i["path"] for i in items if i.get("status") in {"error", "no_pdf_tool", "scanned"}],
        "summary_exists": sp["summary"].exists(),
        "mock_questions": (sp["mock"] / "questions.json").exists(),
        "cheatsheet_exists": sp["cheatsheet_pdf"].exists(),
        "mastery": counts,
        "weakest": [f"{r['id']} {r['name']}" for r in rows[:3]],
    }
    print(json.dumps(info, ensure_ascii=False, indent=2))


def cmd_ingest(args):
    root = Path(args.dir).expanduser().resolve()
    state = load_state(root)
    sp = session_paths(root)
    out_dir = Path(args.out).expanduser().resolve() if args.out else sp["extracted"]
    out_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for raw in args.source:
        p = Path(raw).expanduser().resolve()
        if not p.exists():
            print(f"warning: not found: {p}", file=sys.stderr)
            continue
        if p.is_dir():
            paths += sorted(x for x in p.rglob("*") if x.is_file() and x.suffix.lower() in SUPPORTED)
        elif p.suffix.lower() in SUPPORTED:
            paths.append(p)
        else:
            print(f"warning: unsupported type, skipped: {p}", file=sys.stderr)
    if not paths:
        die("no supported source files found")
    # de-duplicate while preserving order
    seen = set()
    unique = []
    for p in paths:
        key = str(p).lower()
        if key not in seen:
            seen.add(key)
            unique.append(p)
    old = read_json(sp["sources"], {"items": []}) or {"items": []}
    known = {i["path"] for i in old.get("items", [])}
    items = old.get("items", [])
    for p in unique:
        if str(p) in known and not args.force:
            print(f"skip (already ingested): {p}")
            continue
        try:
            item = extract_one(p, out_dir)
        except Exception as exc:
            item = {"path": str(p), "kind": p.suffix.lstrip("."), "pages": None, "chars": 0,
                    "text_file": "", "status": "error", "error": str(exc), "extracted_at": now()}
        if item is None:
            print(f"warning: cannot extract: {p}", file=sys.stderr)
            continue
        items = [i for i in items if i.get("path") != item["path"]] + [item]
        known.add(item["path"])
        pages = item.get("pages")
        unit = "slides" if item.get("kind") == "pptx" else ("pages" if pages else "")
        print(f"{item['status']:12} {item['kind']:5} {str(pages or '-'):>4} {unit:5} "
              f"{item['chars']:>7} chars  {p.name}")
    payload = {"generated_at": now(), "items": items}
    write_json(sp["sources"], payload)
    if state.get("stage") == "intake":
        state["stage"] = "ingest"
        save_state(root, state)
    ok = sum(1 for i in items if i.get("status") == "ok")
    print(f"\n{len(unique)} source(s) processed; {ok}/{len(items)} with usable text.")
    print(f"inventory: {sp['sources']}")
    print(f"extracted: {out_dir}")
    alt = [i for i in items if i.get("status") in {"scanned", "no_pdf_tool", "low_text"}]
    if alt:
        print("\nNeeds vision/office fallback:")
        for i in alt:
            print(f"- {i['path']} ({i['status']})")


def cmd_coverage(args):
    root = Path(args.dir).expanduser().resolve()
    state = load_state(root)
    md_path = Path(args.md).expanduser().resolve() if args.md else session_paths(root)["cheatsheet_md"]
    if not md_path.exists():
        tex = root / "cheatsheet.tex"
        if tex.exists():
            md_path = tex
        else:
            die(f"cheatsheet source not found: tried {md_path} and {tex}")
    text = md_path.read_text(encoding="utf-8-sig")
    norm_text = norm_name(text)
    missing = []
    for t in state.get("topics", []):
        tid = str(t.get("id", "")).strip().lower()
        name = norm_name(str(t.get("name", "")))
        id_hit = bool(tid) and re.search(rf"(?<![a-z0-9]){re.escape(tid)}(?![a-z0-9])", text.lower())
        name_hit = bool(name) and name in norm_text
        if not (id_hit or name_hit):
            missing.append(f"{t['id']} {t['name']}")
    print(json.dumps({"topics": len(state.get("topics", [])), "missing": missing,
                      "covered": len(state.get("topics", [])) - len(missing)},
                     ensure_ascii=False, indent=2))
    if missing:
        raise SystemExit(3)


# --------------------------------------------------------------------------- #
# argparse
# --------------------------------------------------------------------------- #
def build_parser():
    ap = argparse.ArgumentParser(description=__doc__)
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("init", help="create a study session")
    p.add_argument("--dir", required=True)
    p.add_argument("--course", required=True)
    p.add_argument("--exam", default="Final")
    p.add_argument("--date", default="")
    p.add_argument("--language", default="auto")
    p.add_argument("--topics", help="optional topics JSON")
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_init)

    p = sub.add_parser("add-topics", help="merge topics from JSON")
    p.add_argument("--dir", required=True)
    p.add_argument("--json", required=True)
    p.set_defaults(func=cmd_add_topics)

    p = sub.add_parser("stage", help="set workflow stage")
    p.add_argument("--dir", required=True)
    p.add_argument("--set", required=True)
    p.set_defaults(func=cmd_stage)

    p = sub.add_parser("record", help="record one mock-exam attempt")
    p.add_argument("--dir", required=True)
    p.add_argument("--topic", required=True)
    p.add_argument("--qid", default="")
    p.add_argument("--type", default="unknown")
    p.add_argument("--difficulty", type=int, default=3)
    p.add_argument("--correct", choices=["yes", "no", "partial", "skip", "ungraded"], default="ungraded")
    p.add_argument("--score", type=float, default=None)
    p.add_argument("--confidence", type=int, default=None)
    p.add_argument("--seconds", type=int, default=None)
    p.add_argument("--notes", default="")
    p.set_defaults(func=cmd_record)

    p = sub.add_parser("mastery", help="compute mastery and space allocation")
    p.add_argument("--dir", required=True)
    p.add_argument("--write", action="store_true")
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=cmd_mastery)

    p = sub.add_parser("status", help="show session status")
    p.add_argument("--dir", required=True)
    p.set_defaults(func=cmd_status)

    p = sub.add_parser("ingest", help="extract text from lecture/homework files")
    p.add_argument("--dir", required=True)
    p.add_argument("--source", nargs="+", required=True)
    p.add_argument("--out", default="")
    p.add_argument("--force", action="store_true")
    p.set_defaults(func=cmd_ingest)

    p = sub.add_parser("coverage", help="check every state topic appears in cheatsheet.md")
    p.add_argument("--dir", required=True)
    p.add_argument("--md", default="")
    p.set_defaults(func=cmd_coverage)
    return ap


def main(argv=None):
    args = build_parser().parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
