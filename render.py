"""Slide → PNG via LibreOffice → PDF → pdftoppm. Q8: fail loud, no fallbacks."""

from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path


SOFFICE = os.environ.get("PPT2MD_SOFFICE") or shutil.which("soffice") or "/opt/homebrew/bin/soffice"
PDFTOPPM = os.environ.get("PPT2MD_PDFTOPPM") or shutil.which("pdftoppm") or "/opt/homebrew/bin/pdftoppm"


def render(pptx: Path, debug_dir: Path, dpi: int = 150) -> list[Path]:
    """Render a pptx to one PNG per slide in debug_dir, named slide-{N}.png (1-indexed).

    Returns the list of PNG paths in slide order.

    Raises RuntimeError on any failure with stderr surfaced.
    """
    pptx = Path(pptx).resolve()
    debug_dir = Path(debug_dir).resolve()
    debug_dir.mkdir(parents=True, exist_ok=True)

    if not pptx.is_file():
        raise FileNotFoundError(f"pptx not found: {pptx}")
    if not Path(SOFFICE).exists():
        raise RuntimeError(f"LibreOffice not found at {SOFFICE}")
    if not Path(PDFTOPPM).exists():
        raise RuntimeError(f"pdftoppm not found at {PDFTOPPM}")

    with tempfile.TemporaryDirectory(prefix="ppt2md-render-") as td:
        tmp = Path(td)
        # 1. soffice → pdf
        proc = subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "pdf", "--outdir", str(tmp), str(pptx)],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            raise RuntimeError(
                f"soffice failed (exit {proc.returncode}):\nstdout: {proc.stdout}\nstderr: {proc.stderr}"
            )
        pdf = tmp / f"{pptx.stem}.pdf"
        if not pdf.exists() or pdf.stat().st_size == 0:
            raise RuntimeError(f"soffice produced no PDF at {pdf}")

        # 2. pdftoppm → png(s). Output prefix: slide → slide-1.png, slide-2.png, ...
        prefix = debug_dir / "slide"
        # Clean stale slide-*.png
        for stale in debug_dir.glob("slide-*.png"):
            stale.unlink()
        proc = subprocess.run(
            [PDFTOPPM, "-r", str(dpi), "-png", str(pdf), str(prefix)],
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0:
            raise RuntimeError(
                f"pdftoppm failed (exit {proc.returncode}):\nstdout: {proc.stdout}\nstderr: {proc.stderr}"
            )

    pngs = sorted(
        debug_dir.glob("slide-*.png"),
        key=lambda p: int(p.stem.split("-")[-1]),
    )
    if not pngs:
        raise RuntimeError(f"pdftoppm produced no PNGs in {debug_dir}")
    return pngs


if __name__ == "__main__":
    import sys

    if len(sys.argv) != 3:
        print("usage: python -m ppt2md.render <pptx> <out_dir>", file=sys.stderr)
        sys.exit(2)
    out = render(Path(sys.argv[1]), Path(sys.argv[2]))
    for p in out:
        print(p)
