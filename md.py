"""Pure JSON-blocks → Markdown renderer (Q4).

The renderer is deterministic and has no I/O. The VLM emits a SlideDoc dict
matching the schema below; this module turns it into a markdown string.

SlideDoc top-level
------------------
- slide_title: str
- slide_subtitle: str | None
- blocks: list[Block]
- dropped_chrome: list[str]            # manifest shape IDs (for audit)
- confidentiality_marks: list[str]     # rendered as blockquotes near the top
- speaker_notes_used: bool

Block kinds
-----------
- {"kind": "heading", "level": int, "text": str}            # level 1 = biggest on slide
- {"kind": "paragraph", "runs": [{"text": str, "bold"?: bool, "italic"?: bool, "code"?: bool}]}
- {"kind": "image", "manifest_shape_id": str, "alt": str,
       "caption"?: str, "callout"?: {"text": str, "style"?: str}}
- {"kind": "list", "ordered": bool, "items": list[str]}
- {"kind": "table", "rows": list[list[str]]}                # row 0 is header
- {"kind": "quote", "text": str}
- {"kind": "divider"}

Each block helper returns a list of *paragraph strings* (each may contain internal
newlines). The top-level joins paragraphs with a single blank line between them.
"""

from __future__ import annotations

from typing import Any


def _resolve_image_url(image_dir_rel: str, image_path: str) -> str:
    """Return the URL/path used in the markdown image reference.

    If `image_path` is already a URL (contains `://`, e.g. `https://…` or
    `s3://…` returned by an Uploader), pass it through unchanged. Otherwise
    join it with `image_dir_rel` so links resolve from wherever the .md lives.
    """
    if "://" in image_path:
        return image_path
    if image_dir_rel and image_dir_rel != ".":
        return f"{image_dir_rel.rstrip('/')}/{image_path}"
    return image_path


def _render_runs(runs: list[dict[str, Any]]) -> str:
    parts: list[str] = []
    for r in runs:
        t = r.get("text", "")
        if not t:
            continue
        if r.get("code"):
            t = f"`{t}`"
        if r.get("bold") and r.get("italic"):
            t = f"***{t}***"
        elif r.get("bold"):
            t = f"**{t}**"
        elif r.get("italic"):
            t = f"*{t}*"
        parts.append(t)
    return "".join(parts)


def _render_heading(b: dict[str, Any], slide_offset: int) -> list[str]:
    level = max(1, int(b.get("level", 1)))
    n_hash = min(6, level + slide_offset)
    return [f"{'#' * n_hash} {b.get('text', '')}"]


def _render_paragraph(b: dict[str, Any]) -> list[str]:
    runs = b.get("runs")
    if runs:
        text = _render_runs(runs)
    else:
        # Some models emit `text` directly instead of `runs`; accept it.
        text = b.get("text", "")
    return [text] if text else []


def _resolve_image(manifest_shape_id: str, manifest: dict[str, Any]) -> dict[str, Any] | None:
    for s in manifest.get("shapes", []):
        if s["id"] == manifest_shape_id:
            return s
    return None


def _render_image(b: dict[str, Any], manifest: dict[str, Any], image_dir_rel: str) -> list[str]:
    sid = b.get("manifest_shape_id", "")
    shape = _resolve_image(sid, manifest)
    if shape is None or not shape.get("image_path"):
        return [f"<!-- missing image: {sid} -->"]

    rel = _resolve_image_url(image_dir_rel, shape["image_path"])

    def _single_line(s: str) -> str:
        # Markdown alt and italic captions can't span newlines without breaking layout.
        # The validator preserves substring fidelity; the renderer flattens for display.
        return " — ".join(part.strip() for part in s.splitlines() if part.strip())

    paragraphs: list[str] = []
    callout = b.get("callout")
    if callout and callout.get("text"):
        paragraphs.append(f"> {_single_line(callout['text'])}")
    alt = _single_line(b.get("alt") or "")
    paragraphs.append(f"![{alt}]({rel})")
    caption = b.get("caption")
    if caption:
        paragraphs.append(f"*{_single_line(caption)}*")
    return paragraphs


def _render_image_row(b: dict[str, Any], manifest: dict[str, Any], image_dir_rel: str) -> list[str]:
    def _single_line(s: str) -> str:
        return " — ".join(part.strip() for part in s.splitlines() if part.strip())

    cells: list[str] = []
    for item in b.get("images", []) or []:
        sid = item.get("manifest_shape_id", "")
        shape = _resolve_image(sid, manifest)
        if shape is None or not shape.get("image_path"):
            cells.append(f"<td><!-- missing image: {sid} --></td>")
            continue
        rel = _resolve_image_url(image_dir_rel, shape["image_path"])
        parts: list[str] = []
        callout = item.get("callout")
        if callout and callout.get("text"):
            parts.append(f"<blockquote>{_single_line(callout['text'])}</blockquote>")
        alt = _single_line(item.get("alt") or "")
        parts.append(f'<img src="{rel}" alt="{alt}"/>')
        caption = item.get("caption")
        if caption:
            parts.append(f"<br/><em>{_single_line(caption)}</em>")
        cells.append("<td>" + "".join(parts) + "</td>")
    if not cells:
        return []
    paragraphs = ["<table><tr>" + "".join(cells) + "</tr></table>"]
    row_caption = b.get("caption")
    if row_caption:
        paragraphs.append(f"*{_single_line(row_caption)}*")
    return paragraphs


def _render_list(b: dict[str, Any]) -> list[str]:
    ordered = bool(b.get("ordered", False))
    items = b.get("items", [])
    if not items:
        return []
    lines = [f"{i + 1}. {item}" for i, item in enumerate(items)] if ordered else [f"- {item}" for item in items]
    return ["\n".join(lines)]


def _render_table(b: dict[str, Any], manifest: dict[str, Any], image_dir_rel: str) -> list[str]:
    rows = b.get("rows", [])
    if not rows:
        return []

    def _flatten(s: str) -> str:
        return " — ".join(part.strip() for part in s.splitlines() if part.strip())

    def _has_image_cell(rs: list[list[Any]]) -> bool:
        for r in rs:
            for c in r:
                if isinstance(c, dict) and c.get("manifest_shape_id"):
                    return True
        return False

    if not _has_image_cell(rows):
        # Pure-text grid: emit GitHub-flavored pipe-table.
        def _cell(c: Any) -> str:
            text = c if isinstance(c, str) else (c.get("text") or "" if isinstance(c, dict) else "")
            return _flatten(text).replace("|", "\\|")

        lines: list[str] = []
        header = [_cell(c) for c in rows[0]]
        lines.append("| " + " | ".join(header) + " |")
        lines.append("|" + "|".join(["---"] * len(header)) + "|")
        for r in rows[1:]:
            cells = [_cell(c) for c in r] + [""] * (len(header) - len(r))
            lines.append("| " + " | ".join(cells[: len(header)]) + " |")
        return ["\n".join(lines)]

    # Image-bearing grid: emit a single HTML <table> spanning all rows so the
    # markdown renders as one continuous table (markdown pipe-tables can hold
    # inline <img>, but for grids whose cells overlay images on labels, an
    # explicit <table> with <thead>/<tbody> is more reliable across renderers).
    n_cols = max(len(r) for r in rows)

    def _render_cell(c: Any, tag: str) -> str:
        if isinstance(c, str):
            return f"<{tag}>{_flatten(c)}</{tag}>"
        if not isinstance(c, dict):
            return f"<{tag}></{tag}>"
        parts: list[str] = []
        sid = c.get("manifest_shape_id")
        if sid:
            shape = _resolve_image(sid, manifest)
            if shape and shape.get("image_path"):
                rel = _resolve_image_url(image_dir_rel, shape["image_path"])
                alt = _flatten(c.get("alt") or c.get("text") or "")
                parts.append(f'<img src="{rel}" alt="{alt}"/>')
            else:
                parts.append(f"<!-- missing image: {sid} -->")
        text = _flatten(c.get("text") or "")
        if text:
            parts.append((" " if parts else "") + text)
        caption = _flatten(c.get("caption") or "")
        if caption:
            parts.append(f"<br/><em>{caption}</em>")
        return f"<{tag}>{''.join(parts)}</{tag}>"

    out = ["<table>"]
    header_row = rows[0] + [""] * (n_cols - len(rows[0]))
    out.append("<thead><tr>" + "".join(_render_cell(c, "th") for c in header_row) + "</tr></thead>")
    if len(rows) > 1:
        out.append("<tbody>")
        for r in rows[1:]:
            r_padded = list(r) + [""] * (n_cols - len(r))
            out.append("<tr>" + "".join(_render_cell(c, "td") for c in r_padded) + "</tr>")
        out.append("</tbody>")
    out.append("</table>")
    return ["\n".join(out)]


def _render_quote(b: dict[str, Any]) -> list[str]:
    text = b.get("text", "")
    return [f"> {text}"] if text else []


def _render_divider(_: dict[str, Any]) -> list[str]:
    return ["---"]


_BLOCK_RENDERERS = {
    "paragraph": _render_paragraph,
    "image": _render_image,
    "image_row": _render_image_row,
    "list": _render_list,
    "table": _render_table,
    "quote": _render_quote,
    "divider": _render_divider,
}


def render_md(
    slide_doc: dict[str, Any],
    manifest: dict[str, Any],
    image_dir_rel: str,
    slide_number: int | None = None,
    full_slide_image_url: str | None = None,
) -> str:
    """Render one slide_doc to markdown.

    `image_dir_rel` is prepended to manifest `image_path` so links resolve from
    wherever <stem>.md lives. Pass "<stem>" if .md sits alongside the <stem>/
    media folder; pass "" if .md sits inside <stem>/. Image paths that already
    look like URLs (`scheme://…`) are passed through unmodified — useful when
    an Uploader has rewritten the manifest to point at remote storage.

    `slide_number`, when given, is appended to the title as `(slide #N)`.

    `full_slide_image_url`, when given, is used verbatim as the reference image
    URL under the title. Otherwise the local `<image_dir_rel>/media/slide{N}-full.png`
    is constructed.
    """
    paragraphs: list[str] = []

    title = slide_doc.get("slide_title") or ""
    if title:
        if slide_number is not None:
            paragraphs.append(f"# {title} (slide #{slide_number})")
        else:
            paragraphs.append(f"# {title}")

    subtitle = slide_doc.get("slide_subtitle")
    if subtitle:
        paragraphs.append(f"*{subtitle}*")

    marks = slide_doc.get("confidentiality_marks") or []
    if marks:
        paragraphs.append("\n".join(f"> {m}" for m in marks))

    # Reference image directly under the title so the reader can see the source
    # slide before reading the structured extraction below.
    if slide_number is not None:
        if full_slide_image_url is not None:
            full_path = full_slide_image_url
        else:
            full_path = (
                f"{image_dir_rel.rstrip('/')}/media/slide{slide_number}-full.png"
                if image_dir_rel and image_dir_rel != "."
                else f"media/slide{slide_number}-full.png"
            )
        paragraphs.append(
            "Original rendered slide shown below for reference; the structured "
            "extraction follows after."
        )
        paragraphs.append(f"![slide #{slide_number}]({full_path})")

    slide_offset = 1  # in-body level 1 → ##

    blocks = list(slide_doc.get("blocks", []))
    # Defensive dedup: some models re-emit the slide_title as a body H1. Drop it.
    if (
        blocks
        and blocks[0].get("kind") == "heading"
        and (blocks[0].get("text") or "").strip().casefold() == (title or "").strip().casefold()
    ):
        blocks = blocks[1:]

    for b in blocks:
        kind = b.get("kind")
        if kind == "heading":
            paragraphs.extend(_render_heading(b, slide_offset))
        elif kind in _BLOCK_RENDERERS:
            renderer = _BLOCK_RENDERERS[kind]
            if kind in ("image", "image_row", "table"):
                paragraphs.extend(renderer(b, manifest, image_dir_rel))
            else:
                paragraphs.extend(renderer(b))
        else:
            paragraphs.append(f"<!-- unknown block kind: {kind} -->")

    return "\n\n".join(paragraphs) + "\n"
