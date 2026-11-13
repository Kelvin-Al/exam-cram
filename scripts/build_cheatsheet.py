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
    """Replace the body font size and return (patched, marker).

    The template marks the size with a line like
        \\AtBeginDocument{\\fontsize{7pt}{8.2pt}\\selectfont}
    We rewrite that line to keep the template valid, and INSERT an active
    \\fontsize after \\begin{document} (AtBeginDocument fires before multicol
    resets sizes, so the insert is what actually changes the body size).
    """
    new = f"{size_pt}pt" if size_pt % 1 == 0 else f"{size_pt:.1f}pt"
    lines = tex.splitlines(keepends=True)
    out = []
    found = False
    body_marker = "\\begin{document}"
    inserted = False
    for ln in lines:
        if "AtBeginDocument" in ln and "fontsize" in ln:
            out.append(re.sub(r"\d+(?:\.\d+)?pt", new, ln, count=1))
            found = True
            continue
        if not inserted and body_marker in ln:
            indent = " " * (len(ln) - len(ln.lstrip()))
            out.append(ln)
            out.append(f"{indent}\\fontsize{{{new}}}{{{size_pt * 1.15:.1f}pt}}\\selectfont\n")
            inserted = True
            continue
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
    """Rough fill per page = fraction of non-white pixels (density).

    For a 3-column multicol layout, 'content-bottom' is meaningless (columns
    balance into the page top region). Pixel density is a reliable proxy for
    how full the page looks.
    """
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
            dark = 0
            n = 0
            for y in range(0, h, 2):
                for x in range(0, w, 2):
                    if pix[x, y] < 200:
                        dark += 1
                    n += 1
            out.append({"page": p.name, "fill": round(dark / n, 3)})
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
    ap.add_argument("--min-size", type=float, default=5.0)
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
    results = []
    try:
        for size in [round(x, 1) for x in
                     np_arange_desc(args.max_size, args.min_size, args.step)]:
            try:
                compile_tex(tectonic, tex_path, out_pdf, workdir, size, cache)
            except Exception as exc:
                print(f"  size {size:.1f}: {exc}", file=sys.stderr)
                continue
            pages = None
            fill_here = []
            if node and office_cli and out_pdf.exists():
                try:
                    size_prev = workdir / f"preview-{size:.1f}"
                    manifest = render_pdf(out_pdf, size_prev, node, office_cli)
                    pages = int(manifest.get("pageCount", 0))
                    fill_here = measure_fill(size_prev)
                except Exception:
                    pages = None
            max_fill = max((f["fill"] for f in fill_here), default=0.0)
            print(f"size {size:.1f}pt -> {pages} page(s), max_fill={max_fill:.2f}")
            if pages:
                results.append({"size": size, "pages": pages, "fill": max_fill})
        if not results:
            raise SystemExit("no size produced a PDF")

        # Choose result: prefer exactly 2 pages with the largest decided font
        # size that still fits (largest size => most readable). If content
        # underfills even at max size, prefer a single fully-filled page over
        # two nearly-empty pages.
        two_ok = [r for r in results if r["pages"] == 2]
        if two_ok:
            # Keep the LARGEST readable size that still fits 2 pages. When
            # content genuinely underfills (even 12pt gives 2 sparse pages),
            # report ok (content size is what it is); the sparse note explains.
            best = max(two_ok, key=lambda r: (r["fill"] >= 0.02, r["size"]))
        else:
            one_ok = [r for r in results if r["pages"] == 1]
            if one_ok:
                best = max(one_ok, key=lambda r: r["fill"])
            else:
                best = min(results, key=lambda r: r["pages"])
        final_size = best["size"]
        pages = best["pages"]
        max_fill = best["fill"]
        # Rebuild at the chosen size so the committed PDF matches
        compile_tex(tectonic, tex_path, out_pdf, workdir, final_size, cache)
    finally:
        shutil.rmtree(workdir, ignore_errors=True)

    fill = []
    if node and office_cli:
        manifest = render_pdf(final_pdf := out_pdf, prev_dir, node, office_cli)
        pages = int(manifest.get("pageCount", 0))
        fill = measure_fill(prev_dir)
        max_fill = max((f["fill"] for f in fill), default=0.0)
    if pages == 2 and max_fill >= 0.02:
        status = "ok"
    elif pages == 1 and max_fill >= 0.03:
        status = "ok_single"
    elif pages and pages > 2:
        status = "overflow"
    else:
        status = "sparse"
    report = {
        "status": status,
        "pages": pages,
        "size_pt": final_size,
        "pdf": str(out_pdf),
        "preview_dir": str(prev_dir),
        "preview_images": [str(p) for p in sorted(prev_dir.glob("page-*.png"))],
        "page_fill": fill,
        "coverage": cov,
        "note": ("exactly 2 A4 sides, well filled" if status == "ok" else
                 "1 page, well filled (content too small for a full 2-page sheet)" if status == "ok_single" else
                 "2 pages but under-filled — write MORE exam-grade content per topic "
                 "(formulas, conditions, worked patterns), then rebuild" if status == "sparse" else
                 "too much content for 2 pages at the minimum size; trim or drop low-yield items"
                 if status == "overflow" else
                 "content underfills; add more detail"),
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