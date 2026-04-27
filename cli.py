"""Pipeline orchestrator (Q9).

    python -m ppt2md <pptx> <out_dir>

Produces:
    <out_dir>/<stem>.md
    <out_dir>/<stem>/media/*.png
    <out_dir>/<stem>.debug/
        slide-{N}.png, manifest-{N}.json, notes-{N}.txt, chrome-{N}.json,
        slide_doc-{N}.json, chrome_dropped.md
        validation_errors.json   (only on validation failure)

Errors → stderr + nonzero exit. Re-runs overwrite outputs (no caching in v1).
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

from .extract import crop_bbox, extract
from .log import RunLogger
from .md import render_md
from .render import render
from .vlm import ValidationFailure, call_vlm


DEFAULT_VLM_CONCURRENCY = 8


def _write_chrome_audit(
    debug_dir: Path,
    chrome_per_slide: list[dict[str, Any]],
    dropped_per_slide: list[list[str]],
    manifests: list[dict[str, Any]],
) -> None:
    """Audit trail of every shape that was dropped, with reason."""
    lines = ["# Chrome dropped\n"]
    for slide_idx, (chrome, model_dropped, manifest) in enumerate(
        zip(chrome_per_slide, dropped_per_slide, manifests), start=1
    ):
        by_id = {s["id"]: s for s in manifest.get("shapes", [])}
        rows: list[tuple[str, str, str]] = []
        for sid in chrome.get("high", []):
            text = (by_id.get(sid, {}).get("text") or "").replace("\n", " ⏎ ")
            rows.append((sid, "master-inherited", text))
        for sid in chrome.get("medium", []):
            text = (by_id.get(sid, {}).get("text") or "").replace("\n", " ⏎ ")
            rows.append((sid, "repeats across slides", text))
        # Model-flagged: the VLM put the id in dropped_chrome.
        already = {sid for sid, _, _ in rows}
        for sid in model_dropped:
            if sid in already:
                continue
            text = (by_id.get(sid, {}).get("text") or "").replace("\n", " ⏎ ")
            rows.append((sid, "model judged chrome", text))
        if not rows:
            lines.append(f"## Slide {slide_idx}\n\n(none)\n")
            continue
        lines.append(f"## Slide {slide_idx}\n")
        lines.append("| id | reason | text |")
        lines.append("|---|---|---|")
        for sid, reason, text in rows:
            lines.append(f"| `{sid}` | {reason} | {text} |")
        lines.append("")
    (debug_dir / "chrome_dropped.md").write_text("\n".join(lines))


def main(argv: list[str]) -> int:
    if len(argv) != 3:
        print("usage: python -m ppt2md <pptx> <out_dir>", file=sys.stderr, flush=True)
        return 2
    pptx = Path(argv[1])
    out_dir = Path(argv[2])
    if not pptx.is_file():
        print(f"error: pptx not found: {pptx}", file=sys.stderr, flush=True)
        return 1

    stem = pptx.stem
    out_dir.mkdir(parents=True, exist_ok=True)
    debug_dir = out_dir / f"{stem}.debug"
    out_root = out_dir / stem  # holds media/
    debug_dir.mkdir(parents=True, exist_ok=True)
    log = RunLogger(debug_dir / "run.log")
    t_pipeline = time.monotonic()
    log.event("pipeline.start", f"ppt2md: {pptx.name} → {out_dir}", pptx=str(pptx), out_dir=str(out_dir))

    # 1. render
    t = time.monotonic()
    log.event("phase.render.start", f"[1/5] rendering {pptx.name} → PNGs")
    pngs = render(pptx, debug_dir)
    log.event(
        "phase.render.done",
        f"      {len(pngs)} slide(s) rendered ({time.monotonic() - t:.1f}s)",
        slides=len(pngs),
        seconds=round(time.monotonic() - t, 3),
    )

    # 2. extract
    t = time.monotonic()
    log.event("phase.extract.start", "[2/5] extracting manifests + cropping media")
    artifacts = extract(pptx, pngs, out_root, debug_dir)
    n_crops = sum(len(a.image_paths) for a in artifacts)
    log.event(
        "phase.extract.done",
        f"      {n_crops} image-like crop(s) ({time.monotonic() - t:.1f}s)",
        crops=n_crops,
        seconds=round(time.monotonic() - t, 3),
    )

    # 3. VLM call per slide
    t = time.monotonic()
    try:
        concurrency = int(os.environ.get("PPT2MD_VLM_CONCURRENCY", DEFAULT_VLM_CONCURRENCY))
    except ValueError:
        concurrency = DEFAULT_VLM_CONCURRENCY
    concurrency = max(1, min(concurrency, len(artifacts) or 1))
    log.event(
        "phase.vlm.start",
        f"[3/5] calling VLM per slide (concurrency={concurrency})",
        concurrency=concurrency,
    )

    # Pre-load per-slide inputs in slide order.
    inputs_by_idx: dict[int, tuple[Any, ...]] = {}
    for art in artifacts:
        manifest = json.loads(art.manifest_path.read_text())
        notes = art.notes_path.read_text()
        chrome = json.loads(art.chrome_path.read_text())
        png = pngs[art.slide_index - 1]
        inputs_by_idx[art.slide_index] = (manifest, notes, chrome, png)

    docs_by_idx: dict[int, dict[str, Any]] = {}
    attempts_by_idx: dict[int, int] = {}

    def _run_one(slide_index: int) -> tuple[int, dict[str, Any], int]:
        manifest, notes, chrome, png = inputs_by_idx[slide_index]
        log.event("vlm.start", f"      slide {slide_index} → VLM ...", slide_index=slide_index)
        doc, n_attempts = call_vlm(
            png, manifest, notes, chrome,
            debug_dir=debug_dir, slide_index=slide_index, logger=log,
        )
        (debug_dir / f"slide_doc-{slide_index}.json").write_text(
            json.dumps(doc, indent=2, ensure_ascii=False)
        )
        return slide_index, doc, n_attempts

    try:
        with ThreadPoolExecutor(max_workers=concurrency) as pool:
            futures = {pool.submit(_run_one, art.slide_index): art.slide_index for art in artifacts}
            for fut in as_completed(futures):
                idx, doc, n_attempts = fut.result()
                docs_by_idx[idx] = doc
                attempts_by_idx[idx] = n_attempts
    except ValidationFailure as e:
        out = [v.to_dict() for v in e.violations]
        (debug_dir / "validation_errors.json").write_text(
            json.dumps(out, indent=2, ensure_ascii=False)
        )
        log.event(
            "phase.vlm.failed",
            f"error: validation failed after retries; see {debug_dir / 'validation_errors.json'}",
            n_violations=len(e.violations),
        )
        return 1

    # Re-assemble in slide order so downstream stages stay deterministic.
    slide_docs: list[dict[str, Any]] = [docs_by_idx[art.slide_index] for art in artifacts]
    manifests: list[dict[str, Any]] = [inputs_by_idx[art.slide_index][0] for art in artifacts]
    chrome_per_slide: list[dict[str, Any]] = [inputs_by_idx[art.slide_index][2] for art in artifacts]
    # Attempts histogram — generalization monitor. A healthy run is mostly
    # bucket "1×". A spike in 2×/3× across many slides on a new deck is the
    # earliest signal that the validator is over-fit (or the model is drifting).
    attempts_list = [attempts_by_idx[art.slide_index] for art in artifacts]
    hist: dict[int, int] = {}
    for a in attempts_list:
        hist[a] = hist.get(a, 0) + 1
    n = len(attempts_list)
    avg = sum(attempts_list) / n if n else 0.0
    hist_pretty = ", ".join(f"{k}×: {v}" for k, v in sorted(hist.items()))
    log.event(
        "phase.vlm.done",
        f"      VLM phase: {time.monotonic() - t:.1f}s for {len(slide_docs)} slide(s); "
        f"attempts avg={avg:.2f} hist=[{hist_pretty}]",
        seconds=round(time.monotonic() - t, 3),
        slides=len(slide_docs),
        attempts_avg=round(avg, 3),
        attempts_hist=hist,
        attempts_per_slide=attempts_by_idx,
    )

    # 3.5 Lazy crop: the model may reference Group shapes that extract.py skipped
    # because their bbox was below GROUP_AREA_THRESHOLD. Crop those on demand so
    # md.py never has to render `<!-- missing image: ... -->`.
    media_dir = out_root / "media"
    media_dir.mkdir(parents=True, exist_ok=True)
    n_lazy = 0
    for art, doc, manifest in zip(artifacts, slide_docs, manifests):
        by_id = {s["id"]: s for s in manifest.get("shapes", [])}
        existing = sum(1 for s in manifest.get("shapes", []) if s.get("image_path"))
        fig_idx = existing  # continue numbering from where extract.py stopped

        def _ensure_cropped(sid: str) -> None:
            nonlocal fig_idx, n_lazy
            shape = by_id.get(sid)
            if shape is None or shape.get("image_path") or not shape.get("bbox_frac"):
                return
            fig_idx += 1
            crop_name = f"slide{art.slide_index}-fig{fig_idx}.png"
            crop_path = media_dir / crop_name
            crop_bbox(pngs[art.slide_index - 1], shape["bbox_frac"], crop_path)
            shape["image_path"] = f"media/{crop_name}"
            n_lazy += 1
            log.event(
                "vlm.lazy_crop",
                f"      slide {art.slide_index}: lazy-cropped {sid} → {crop_name}",
                slide_index=art.slide_index,
                shape_id=sid,
                kind=shape.get("kind"),
                bbox_frac=shape.get("bbox_frac"),
                image_path=shape["image_path"],
            )

        for b in doc.get("blocks", []) or []:
            kind = b.get("kind")
            if kind == "image":
                _ensure_cropped(b.get("manifest_shape_id", ""))
            elif kind == "image_row":
                for it in b.get("images", []) or []:
                    _ensure_cropped(it.get("manifest_shape_id", ""))
            elif kind == "table":
                for row in b.get("rows", []) or []:
                    for cell in row:
                        if isinstance(cell, dict) and cell.get("manifest_shape_id"):
                            _ensure_cropped(cell["manifest_shape_id"])
    if n_lazy:
        log.event(
            "phase.vlm.lazy_crop_done",
            f"      lazy-cropped {n_lazy} referenced shape(s) that were below the area threshold",
            n_lazy=n_lazy,
        )

    # 4. join (v1 N=1: identity; for N>1 we just concatenate per-slide markdown)
    log.event("phase.md.start", "[4/5] rendering markdown")
    parts: list[str] = []
    for art, doc, manifest in zip(artifacts, slide_docs, manifests):
        # Copy the full rendered slide PNG into media/ so the markdown's leading
        # reference image resolves alongside the other figures.
        src = pngs[art.slide_index - 1]
        dst = media_dir / f"slide{art.slide_index}-full.png"
        shutil.copyfile(src, dst)
        parts.append(render_md(doc, manifest, image_dir_rel=stem, slide_number=art.slide_index))
    md_text = "\n".join(parts)
    md_path = out_dir / f"{stem}.md"
    md_path.write_text(md_text)
    log.event("phase.md.done", f"      wrote {md_path} ({len(md_text)} chars)", path=str(md_path), chars=len(md_text))

    # 5. audit trail
    log.event("phase.audit.start", "[5/5] writing chrome audit trail")
    _write_chrome_audit(
        debug_dir,
        chrome_per_slide,
        [doc.get("dropped_chrome", []) or [] for doc in slide_docs],
        manifests,
    )
    log.event(
        "pipeline.done",
        f"done: {md_path}  ({time.monotonic() - t_pipeline:.1f}s total)",
        seconds_total=round(time.monotonic() - t_pipeline, 3),
    )
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
