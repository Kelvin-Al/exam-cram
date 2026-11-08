#!/usr/bin/env python3
"""Build a dense two-sided A4 cheat sheet with LaTeX (tectonic).

This replaces the old DOCX-based pipeline. The model authors `cheatsheet.tex`
from `templates/cheatsheet.tex`; this script:

  1. finds the tectonic engine (auto-detect or --tectonic);
  2. compiles the .tex, adjusting the body font size so the PDF is exactly
     2 A4 pages;
  3. renders PNG previews with the bundled LibreOffice Kit;
  4. checks that every state topic id appears in the .tex source.

Usage:
    python build_cheatsheet.py --tex cheatsheet.tex --out cheatsheet.pdf \
        --state state.json --json build-report.json
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

FONT_LINE_RE = re.compile(
    r"(\\AtBeginDocument\{)\\fontsize\{[^}]*\}?\{[^}]*\}(\\selectfont\})"
)


# --------------------------------------------------------------------------- #
# engine discovery
# --------------------------------------------------------------------------- #
def find_tectonic(explicit: str | None = None) -> Path | None:
    if explicit:
        p = Path(explicit).expanduser()
        if p.exists():
            return p
        raise SystemExit(f"tectonic not found at {explicit}")
    env = os.environ.get("EXAM_CRAM_TECTONIC")
    if env:
        p = Path(env).expanduser()
        if p.exists():
            return p
    which = shutil.which("tectonic")
    if which:
        return Path(which)
    # common install locations
    home = Path.home()
    cands = [
        home / ".local" / "bin" / "tectonic.exe",
        home / ".local" / "bin" / "tectonic",
        home / "bin" / "tectonic.exe",
        Path("_tools") / "tectonic.exe",
        Path("tools") / "tectonic.exe",
    ]
    if os.name == "nt":
        cands += [
            Path(r"C:\texlive\bin\windows\tectonic.exe"),
            Path(r"C:\Users\admin\AppData\Local\Programs\tectonic\tectonic.exe"),
        ]
    else:
        cands += [Path("/usr/local/bin/tectonic"), Path("/opt/homebrew/bin/tectonic"),
                  Path("/usr/bin/tectonic")]
    for c in cands:
        if c.exists():
            return c
    return None


def tectonic_help():
    return (
        "\nLaTeX engine 'tectonic' not found. Install one of:\n"
        "  PowerShell:  iwr https://github.com/tectonic-typesetting/tectonic/releases/\\\n"
        "               latest/download/tectonic-0.17.0-x86_64-pc-windows-msvc.zip -o tect.zip; \\\n"
        "               Expand-Archive tect.zip .; move .\\tectonic.exe $HOME\\.local\\bin\\\n"
        "  or set  $env:EXAM_CRAM_TECTONIC = 'C:\\path\\to\\tectonic.exe'\n"
        "  or install TeX Live / MiKTeX and put tectonic on PATH."
    )


# --------------------------------------------------------------------------- #
# compile + fit
# --------------------------------------------------------------------------- #
def patch_font(tex: str, size_pt: float) -> tuple[str, str]:
    """Replace the AtBeginDocument fontsize line and return (patched, marker)."""
    new = rf"{size_pt}pt" if size_pt % 1 == 0 else f"{size_pt:.1f}pt"
    lines = tex.splitlines(keepends=True)
    out = []
    found = False
    for ln in lines:
        if "AtBeginDocument" in ln and "fontsize" in ln:
            out.append(re.sub(r"\d+(?:\.\d+)?pt", new, ln, count=1))
            found = True
        else:
            out.append(ln)
    if not found:
        raise SystemExit("fontsize line not found in tex template")
    return "".join(out), new


def compile_tex(tectonic: Path, tex_path: Path, out_pdf: Path, workdir: Path,
                size_pt: float, cache_dir: Path | None = None) -> Path:
    tex_src = tex_path.read_text(encoding="utf-8")
    patched, marker = patch_font(tex_src, size_pt)
    stage_tex = workdir / f"stage-{marker.replace('.', 'p')}.tex"
    stage_tex.write_text(patched, encoding="utf-8")
    env = dict(os.environ)
    if cache_dir:
        env["TECTONIC_CACHE_DIR"] = str(cache_dir)
    result = subprocess.run(
        [str(tectonic), str(stage_tex.name)],
        cwd=workdir, capture_output=True, text=True, encoding="utf-8",
        errors="replace", timeout=600, env=env,
    )
    pdf = workdir / (stage_tex.stem + ".pdf")
    if result.returncode != 0 or not pdf.exists():
        err = (result.stderr or result.stdout or "")[-2000:]
        raise RuntimeError(f"tectonic compile failed at {marker}: {err}")
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(pdf, out_pdf)
    return out_pdf, marker


def render_pdf(pdf: Path, out_dir: Path, node: Path, cli: Path) -> dict:
    shutil.rmtree(out_dir, ignore_errors=True)
    res = subprocess.run(
        [str(node), str(cli), "render", "--input", str(pdf), "--output-dir", str(out_dir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
    )
    if res.returncode != 0:
        raise RuntimeError(f"render failed: {res.stderr or res.stdout}")
    return json.loads(res.stdout)


def find_node():
    env = os.environ.get("DSH_NODE_EXECUTABLE")
    if env and Path(env).exists():
        return Path(env)
    home = Path.home()
    hits = sorted(home.glob(".dsh/dsh-runtimes/*/dependencies/node/bin/node.exe")) or \
           sorted(home.glob(".dsh/dsh-runtimes/*/dependencies/node/bin/node"))
    if hits:
        return hits[0]
    return None


def find_office_cli():
    env = os.environ.get("DSH_OFFICE_CLI")
    if env and Path(env).exists():
        return Path(env)
    rel = Path("app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit/lib/cli.js")
    roots = []
    if os.name == "nt":
        local = os.environ.get("LOCALAPPDATA")
        if local:
            roots.append(Path(local) / "Programs" / "DeepSeek Harness" / "resources")
        roots += [Path(r"C:\Program Files\DeepSeek Harness\resources"),
                  Path(r"C:\Program Files (x86)\DeepSeek Harness\resources")]
    elif sys.platform == "darwin":
        roots += [Path("/Applications/DeepSeek Harness.app/Contents/Resources")]
    else:
        roots += [Path("/opt/DeepSeek Harness/resources"), Path("/usr/share/DeepSeek Harness/resources")]
    for root in roots:
        p = root / rel
        if p.exists():
            return p
    which = shutil.which("dsoffice")
    return Path(which) if which else None


# --------------------------------------------------------------------------- #
# coverage
# --------------------------------------------------------------------------- #
def _norm(text: str) -> str:
    return re.sub(r"[^a-z0-9\u4e00-\u9fff]+", " ", (text or "").lower()).strip()


def coverage_check(tex_src: str, state_path: Path | None) -> dict:
    if not state_path:
        return {"checked": False}
    state = json.loads(state_path.read_text(encoding="utf-8-sig"))
    low = tex_src.lower()
    norm = _norm(tex_src)
    missing = []
    for t in state.get("topics", []):
        tid = str(t.get("id", "")).strip().lower()
        name = _norm(str(t.get("name", "")))
        id_hit = bool(tid) and re.search(rf"(?<![a-z0-9]){re.escape(tid)}(?![a-z0-9])", low)
        name_hit = bool(name) and name in norm
        if not (id_hit or name_hit):
            missing.append(f"{t.get('id')} {t.get('name')}")
    return {"checked": True, "topics": len(state.get("topics", [])),
            "covered": len(state.get("topics", [])) - len(missing), "missing": missing}


def measure_fill(prev_dir: Path) -> list[dict]:
    try:
        from PIL import Image
    except Exception:
        return []
    out = []
    for p in sorted(prev_dir.glob("page-*.png")):
        try:
            im = Image.open(p).convert("L")
            w, h = im.size
            pix = im.load()
            last = 0
            for y in range(0, h, 3):
                if any(pix[x, y] < 200 for x in range(0, w, 6)):
                    last = y
            out.append({"page": p.name, "fill": round(last / h, 3)})
        except Exception:
            continue
    return out


# --------------------------------------------------------------------------- #
# main
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tex", required=True, help="cheatsheet.tex source")
    ap.add_argument("--out", required=True, help="final PDF path")
    ap.add_argument("--state", default="")
    ap.add_argument("--json", default="")
    ap.add_argument("--tectonic", default="", help="explicit engine path")
    ap.add_argument("--min-size", type=float, default=5.5)
    ap.add_argument("--max-size", type=float, default=12.0)
    ap.add_argument("--step", type=float, default=0.5)
    args = ap.parse_args(argv)

    tex_path = Path(args.tex).expanduser().resolve()
    if not tex_path.exists():
        raise SystemExit(f"--tex not found: {tex_path}")
    tex_src = tex_path.read_text(encoding="utf-8")
    out_pdf = Path(args.out).expanduser().resolve()
    state_path = Path(args.state).expanduser().resolve() if args.state else None
    prev_dir = out_pdf.with_name(out_pdf.stem + "-preview")

    cov = coverage_check(tex_src, state_path)
    if cov.get("missing"):
        print("coverage warning: " + "; ".join(cov["missing"]), file=sys.stderr)

    tectonic = find_tectonic(args.tectonic or None)
    if not tectonic:
        print(tectonic_help(), file=sys.stderr)
        raise SystemExit(2)
    node = find_node()
    office_cli = find_office_cli()

    workdir = out_pdf.parent / f".tex-build-{os.getpid()}"
    workdir.mkdir(parents=True, exist_ok=True)
    cache = out_pdf.parent / ".tectonic-cache"
    try:
        best = None
        for size in [round(x, 1) for x in
                     np_arange_desc(args.max_size, args.min_size, args.step)]:
            try:
                compile_tex(tectonic, tex_path, out_pdf, workdir, size, cache)
            except Exception as exc:
                print(f"  size {size:.1f}: {exc}", file=sys.stderr)
                continue
            pages = None
            if node and office_cli and out_pdf.exists():
                try:
                    manifest = render_pdf(out_pdf, prev_dir, node, office_cli)
                    pages = int(manifest.get("pageCount", 0))
                except Exception:
                    pages = None
            print(f"size {size:.1f}pt -> {'?' if pages is None else pages} page(s)")
            if pages == 2:
                best = (size, out_pdf)
                break
            if best is None or (pages is not None and pages < 2 and best[0] < size):
                best = (size, out_pdf) if pages is not None and pages < 2 else best
        if best is None:
            raise SystemExit("no size produced a PDF")
        final_size, final_pdf = best
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    pages = None
    fill = []
    if node and office_cli:
        manifest = render_pdf(final_pdf, prev_dir, node, office_cli)
        pages = int(manifest.get("pageCount", 0))
        fill = measure_fill(prev_dir)
    max_fill = max((f["fill"] for f in fill), default=0.0)
    if pages == 2 and max_fill >= 0.7:
        status = "ok"
    elif pages == 2 and max_fill < 0.7:
        status = "sparse"
    elif pages and pages > 2:
        status = "overflow"
    else:
        status = "underfull"
    report = {
        "status": status,
        "pages": pages,
        "size_pt": final_size,
        "pdf": str(final_pdf),
        "preview_dir": str(prev_dir),
        "preview_images": [str(p) for p in sorted(prev_dir.glob("page-*.png"))],
        "page_fill": fill,
        "coverage": cov,
        "note": ("exactly 2 A4 sides, well filled" if status == "ok" else
                 "2 pages but under-filled — write MORE exam-grade content per topic "
                 "(formulas, conditions, worked patterns), then rebuild" if status == "sparse" else
                 "too much content for 2 pages at the minimum size; trim or drop low-yield items"
                 if status == "overflow" else
                 "content underfills 2 pages; add more detail"),
    }
    if status == "sparse":
        report["advice"] = ("Increase content density: each topic should carry 10-25 lines of "
                            "formulas + conditions + traps, mirroring a real cram sheet. "
                            "Rebuild after expanding.")
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                                   encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    if status == "overflow":
        raise SystemExit(4)
    return 0


def np_arange_desc(start, stop, step):
    out = []
    x = start
    while x >= stop - 1e-9:
        out.append(round(x, 1))
        x = round(x - step, 1)
    return out


if __name__ == "__main__":
    sys.exit(main())