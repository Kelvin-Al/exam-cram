#!/usr/bin/env python3
"""Build the multi-page STUDY NOTES PDF (复习材料) with LaTeX (tectonic).

The model authors `study-notes.tex` from `templates/study-notes.tex`; this
script compiles it and reports the page count + renders previews. Unlike the
2-page cheat sheet, study notes are a READABLE document — no page fitting,
no density requirement; just correct, complete revision content.

Usage:
    python build_study_notes.py --tex study-notes.tex --out study-notes.pdf \
        --json build-report.json
"""
from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
from pathlib import Path


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
        roots += [Path("/opt/DeepSeek Harness/resources"),
                  Path("/usr/share/DeepSeek Harness/resources")]
    for root in roots:
        p = root / rel
        if p.exists():
            return p
    which = shutil.which("dsoffice")
    return Path(which) if which else None


def render_pdf(pdf: Path, out_dir: Path, node: Path, cli: Path) -> dict:
    shutil.rmtree(out_dir, ignore_errors=True)
    res = subprocess.run(
        [str(node), str(cli), "render", "--input", str(pdf), "--output-dir", str(out_dir)],
        capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300,
    )
    if res.returncode != 0:
        raise RuntimeError(f"render failed: {res.stderr or res.stdout}")
    return json.loads(res.stdout)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tex", required=True, help="study-notes.tex source")
    ap.add_argument("--out", required=True, help="final PDF path")
    ap.add_argument("--json", default="")
    ap.add_argument("--tectonic", default="")
    args = ap.parse_args(argv)

    tex_path = Path(args.tex).expanduser().resolve()
    if not tex_path.exists():
        raise SystemExit(f"--tex not found: {tex_path}")
    out_pdf = Path(args.out).expanduser().resolve()
    prev_dir = out_pdf.with_name(out_pdf.stem + "-preview")

    tectonic = find_tectonic(args.tectonic or None)
    if not tectonic:
        print(tectonic_help(), file=sys.stderr)
        raise SystemExit(2)
    node = find_node()
    office_cli = find_office_cli()

    cache = out_pdf.parent / ".tectonic-cache"
    env = dict(os.environ)
    if not env.get("TECTONIC_CACHE_DIR"):
        env["TECTONIC_CACHE_DIR"] = str(cache)
    out_pdf.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [str(tectonic), str(tex_path)],
        cwd=tex_path.parent, capture_output=True, text=True,
        encoding="utf-8", errors="replace", timeout=600, env=env,
    )
    produced = tex_path.with_suffix(".pdf")
    if result.returncode != 0 or not produced.exists():
        err = (result.stderr or result.stdout or "")[-2500:]
        raise SystemExit(f"tectonic compile failed: {err}")
    shutil.copyfile(produced, out_pdf)

    pages = None
    preview_images = []
    if node and office_cli:
        try:
            manifest = render_pdf(out_pdf, prev_dir, node, office_cli)
            pages = int(manifest.get("pageCount", 0))
            preview_images = [str(p) for p in sorted(prev_dir.glob("page-*.png"))]
        except Exception as exc:
            print(f"warning: preview render failed: {exc}", file=sys.stderr)

    report = {
        "status": "ok",
        "pages": pages,
        "pdf": str(out_pdf),
        "preview_dir": str(prev_dir),
        "preview_images": preview_images,
        "note": "multi-page study notes PDF rendered (readable revision material, not the 2-page cheat sheet)",
    }
    if args.json:
        Path(args.json).write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n",
                                   encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())