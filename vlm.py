"""VLM call layer (Q7).

Constraints:
- openai SDK with custom base_url (defaults Moonshot Kimi).
- Tool use forces structured output via the `emit_slide_doc` function schema.
- Post-call substring validator: every emitted text must be a substring of some
  manifest shape's `text`; every `manifest_shape_id` must resolve; every image
  ref must point to an image-like shape (Picture/Group/GraphicFrame).
- Up to 2 retries with a structured retry message ({block_index, field, value, reason}).
- On persistent failure, ValidationFailure is raised carrying the violations list
  so the orchestrator can write `validation_errors.json` and exit nonzero.

The system prompt and tool schema are kept byte-stable across calls so Moonshot
auto-caches the prefix (Q7 sub-defaults).
"""

from __future__ import annotations

import base64
import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from openai import APIConnectionError, APITimeoutError, OpenAI


DEFAULT_BASE_URL = "https://api.moonshot.cn/v1"
DEFAULT_MODEL = "kimi-k2.6"


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    try:
        return int(raw)
    except ValueError as e:
        raise RuntimeError(f"invalid {name}={raw!r}: {e}") from e


def _env_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    if raw is None or raw == "":
        return default
    try:
        return float(raw)
    except ValueError as e:
        raise RuntimeError(f"invalid {name}={raw!r}: {e}") from e


# Outer retry budget (validator-driven). Override via PPT2MD_MAX_RETRIES.
MAX_RETRIES = _env_int("PPT2MD_MAX_RETRIES", 4)
# Per-request HTTP timeout in seconds. Override via PPT2MD_REQUEST_TIMEOUT_S.
REQUEST_TIMEOUT_S = _env_float("PPT2MD_REQUEST_TIMEOUT_S", 120.0)
# Disable the OpenAI SDK's internal retry loop — our outer MAX_RETRIES loop
# is the single source of truth. Otherwise a hang costs (SDK_retries+1) ×
# REQUEST_TIMEOUT_S per outer attempt, which compounds badly.
SDK_MAX_RETRIES = 0

IMAGE_LIKE_KINDS = {"Picture", "Group", "GraphicFrame"}


# ---- Tool schema -----------------------------------------------------------

# Compact JSON schema. Moonshot tolerates loose `oneOf`-less variants well; we
# rely on the post-call validator for content fidelity rather than asking the
# model server to enforce per-kind block shapes. Required fields keep the
# essentials anchored.
EMIT_SLIDE_DOC_SCHEMA: dict[str, Any] = {
    "type": "object",
    "additionalProperties": False,
    "properties": {
        "slide_title": {"type": "string"},
        "slide_subtitle": {"type": ["string", "null"]},
        "blocks": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": True,
                "properties": {
                    "kind": {
                        "type": "string",
                        "enum": ["heading", "paragraph", "image", "image_row", "list", "table", "quote", "divider"],
                    },
                    "level": {"type": "integer", "minimum": 1, "maximum": 6},
                    "text": {"type": "string"},
                    "runs": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "text": {"type": "string"},
                                "bold": {"type": "boolean"},
                                "italic": {"type": "boolean"},
                                "code": {"type": "boolean"},
                            },
                            "required": ["text"],
                        },
                    },
                    "manifest_shape_id": {"type": "string"},
                    "alt": {"type": "string"},
                    "caption": {"type": "string"},
                    "callout": {
                        "type": "object",
                        "properties": {"text": {"type": "string"}, "style": {"type": "string"}},
                        "required": ["text"],
                    },
                    "images": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "manifest_shape_id": {"type": "string"},
                                "alt": {"type": "string"},
                                "caption": {"type": "string"},
                                "callout": {
                                    "type": "object",
                                    "properties": {"text": {"type": "string"}, "style": {"type": "string"}},
                                    "required": ["text"],
                                },
                            },
                            "required": ["manifest_shape_id"],
                        },
                    },
                    "ordered": {"type": "boolean"},
                    "items": {"type": "array", "items": {"type": "string"}},
                    "rows": {
                        "type": "array",
                        "items": {
                            "type": "array",
                            "items": {
                                # A cell is either a plain string (text-only grid)
                                # or an object describing an image-bearing cell.
                                # Object cells may carry text alongside the image
                                # for grids whose cells overlay images on labels.
                                "oneOf": [
                                    {"type": "string"},
                                    {
                                        "type": "object",
                                        "additionalProperties": False,
                                        "properties": {
                                            "text": {"type": "string"},
                                            "manifest_shape_id": {"type": "string"},
                                            "alt": {"type": "string"},
                                            "caption": {"type": "string"},
                                        },
                                    },
                                ],
                            },
                        },
                    },
                },
                "required": ["kind"],
            },
        },
        "dropped_chrome": {"type": "array", "items": {"type": "string"}},
        "confidentiality_marks": {"type": "array", "items": {"type": "string"}},
        "speaker_notes_used": {"type": "boolean"},
    },
    "required": [
        "slide_title",
        "blocks",
        "dropped_chrome",
        "confidentiality_marks",
        "speaker_notes_used",
    ],
}

EMIT_SLIDE_DOC_TOOL: dict[str, Any] = {
    "type": "function",
    "function": {
        "name": "emit_slide_doc",
        "description": (
            "Emit the structured representation of one slide. Every text string must be a "
            "substring of some manifest shape's `text` field; every `manifest_shape_id` "
            "must reference an image-like shape (Picture, Group, GraphicFrame). "
            "Use `image_row` when figures sit side-by-side in one visual row; "
            "otherwise use `image`. For a multi-row data grid whose cells overlay "
            "images on labels, keep using a single `table` block and put image cells "
            "as objects `{manifest_shape_id, text?, alt?}` inside `rows`."
        ),
        "parameters": EMIT_SLIDE_DOC_SCHEMA,
    },
}

# Byte-stable system prompt — keep this string immutable to maximize cache hits.
# Each numbered rule is documented (purpose + example) in
# `decisions.md` → "Q7 → System-prompt rule catalog". When you edit a rule here,
# update the catalog in the same commit.
SYSTEM_PROMPT = (
    "You convert a single slide into a structured JSON document by calling the "
    "`emit_slide_doc` function exactly once.\n\n"
    "Inputs you receive:\n"
    "- slide_image: a 150-DPI PNG render of the slide.\n"
    "- manifest: a JSON list of every shape on the slide (id, kind, bbox, text, ...).\n"
    "- notes: the slide's speaker notes (may be empty).\n"
    "- chrome: shape ids the structural pre-pass flagged as likely template chrome.\n\n"
    "Hard rules (a downstream validator will reject violations):\n"
    "1. Every text you emit (slide_title, paragraph runs, list items, table cells, "
    "image alt/caption/callout, quote text, heading text, confidentiality_marks, "
    "subtitle) must appear as a substring of some manifest shape's `text` field. "
    "If there is no good substring for a slot, omit the slot rather than invent text.\n"
    "2. Every `manifest_shape_id` you reference in an image block must be the id of "
    "a manifest shape whose kind is Picture, Group, or GraphicFrame. Never invent ids.\n"
    "3. List `dropped_chrome` for every shape id you treat as template chrome (footer, "
    "page number, watermark, repeated logo). Do NOT include their text in any block.\n"
    "4. Surface confidentiality marks (e.g. 'Confidential') in `confidentiality_marks` "
    "rather than as paragraphs, when present.\n"
    "5. Heading levels are slide-relative: level 1 = the largest heading on this slide.\n"
    "6. Preserve reading order. Bind callouts to their visual target by spatial overlap.\n"
    "7. When two or more figures sit side-by-side in a single visual row, group them in "
    "one `image_row` block with an `images` array (each entry has manifest_shape_id and "
    "optional per-cell alt/caption/callout). The block itself may also have a `caption` "
    "field for an umbrella label that describes the whole row — when a short label sits "
    "visually adjacent to the row (typically above or below), put it in `image_row.caption` "
    "rather than as a separate paragraph block. The row-level caption must itself be a "
    "substring of one single manifest shape's text; do NOT concatenate two shapes' texts "
    "to fabricate a richer caption — pick the better single substring or omit. "
    "Use `image` only for stand-alone figures.\n"
    "8. When the slide shows a horizontal strip of short text labels (e.g. a "
    "product-progression bar, a column-header row, a stage-by-stage timeline), emit "
    "it as a one-row `table` block — each label becomes one cell in `rows[0]`. Do NOT "
    "use `list` or separate paragraphs for horizontal text strips.\n"
    "9. Image-bearing data grid (NARROW rule). Use ONE `table` block with object cells "
    "`{manifest_shape_id, text?, alt?}` ONLY when ALL of the following hold: "
    "(a) every column has the same number of rows, "
    "(b) every data cell is a single short label-with-image (not a bulleted list, "
    "not multiple paragraphs, not multiple figures stacked together), "
    "(c) the cells form a clean rectangular grid with crisp horizontal and vertical "
    "rules visible in the image. "
    "When this rule applies, do NOT split the grid into one `image_row` per row, "
    "and do NOT emit a separate `table` block per row.\n"
    "If a slide shows multiple columns whose contents are lists of bullets, mixed "
    "paragraphs, or multiple figures per column entry, treat each column as its own "
    "section: emit a `heading` for the column label, then `paragraph` / `list` / "
    "`image` / `image_row` blocks for its contents. Do NOT force such a layout into "
    "a single `table` block. `image_row` remains reserved for a single row of "
    "stand-alone figures, not for rows of a data grid."
)


# ---- Validator -------------------------------------------------------------


@dataclass
class Violation:
    block_index: int  # -1 for top-level fields
    field: str
    value: Any
    reason: str

    def to_dict(self) -> dict[str, Any]:
        return {
            "block_index": self.block_index,
            "field": self.field,
            "value": self.value,
            "reason": self.reason,
        }


class ValidationFailure(Exception):
    """Raised after retries are exhausted. Carries the list of violations."""

    def __init__(self, violations: list[Violation]):
        super().__init__(f"validation failed after retries: {len(violations)} violation(s)")
        self.violations = violations


def _normalize_ws(s: str) -> str:
    """Collapse all whitespace runs to a single space and strip ends.

    Manifest text often has double-space artifacts ("Coronus:  Dielectric")
    inherited from PPT layouts; models silently normalize them. A byte-exact
    substring check then rejects the (faithful) transcription. We compare on
    a whitespace-normalized form on both sides instead.
    """
    return " ".join(s.split())


def _is_substring_of_any(s: str, manifest_texts: list[str]) -> bool:
    if not s:
        return True
    if any(s in t for t in manifest_texts):
        return True
    ns = _normalize_ws(s)
    return any(ns in _normalize_ws(t) for t in manifest_texts)


# Fields whose value is the model's free-form description, not a transcription
# of slide text. The substring rule does not apply (alt is an a11y description,
# not OCR; treating it as transcription is a validator concept error).
_DESCRIPTIVE_FIELDS = ("alt",)


def _check_text(value: str, field: str, idx: int, texts: list[str], out: list[Violation]) -> None:
    if not isinstance(value, str):
        return
    # Skip substring check for descriptive fields (alt, …).
    if any(field == d or field.endswith("." + d) for d in _DESCRIPTIVE_FIELDS):
        return
    if value and not _is_substring_of_any(value, texts):
        out.append(
            Violation(
                block_index=idx,
                field=field,
                value=value,
                reason="text is not a substring of any manifest shape's `text`",
            )
        )


def validate(slide_doc: dict[str, Any], manifest: dict[str, Any]) -> list[Violation]:
    """Return a list of violations. Empty list = pass."""
    violations: list[Violation] = []

    shapes = manifest.get("shapes", [])
    texts = [s.get("text", "") for s in shapes if s.get("text")]
    by_id: dict[str, dict[str, Any]] = {s["id"]: s for s in shapes}

    # Reject empty docs when the manifest has substantive content. The model
    # would sometimes call emit_slide_doc({}) as an early-out on dense slides;
    # without this check the validator passes (nothing to flag) and the slide
    # silently renders blank in markdown.
    has_title = bool(slide_doc.get("slide_title"))
    has_subtitle = bool(slide_doc.get("slide_subtitle"))
    has_blocks = bool(slide_doc.get("blocks"))
    if not (has_title or has_subtitle or has_blocks):
        manifest_has_text = bool(texts)
        manifest_has_image = any(s.get("kind") in IMAGE_LIKE_KINDS for s in shapes)
        if manifest_has_text or manifest_has_image:
            violations.append(
                Violation(
                    block_index=-1,
                    field="<root>",
                    value=None,
                    reason=(
                        "slide_doc is empty but the manifest has content "
                        "(text-bearing shapes and/or image-like shapes); "
                        "emit slide_title, slide_subtitle, and/or blocks"
                    ),
                )
            )
            return violations

    # Top-level scalar text fields.
    _check_text(slide_doc.get("slide_title", ""), "slide_title", -1, texts, violations)
    if slide_doc.get("slide_subtitle"):
        _check_text(slide_doc["slide_subtitle"], "slide_subtitle", -1, texts, violations)
    # NOTE: confidentiality_marks is a classification label (e.g. "Confidential",
    # "Internal Use Only") — typically inherited from the master slide and often
    # filtered out as chrome before the manifest is built. We do not enforce the
    # substring check on it; the model's job here is to *categorize* the slide,
    # not to transcribe a string that's guaranteed to be present.

    for idx, b in enumerate(slide_doc.get("blocks", []) or []):
        kind = b.get("kind")
        if kind == "heading":
            _check_text(b.get("text", ""), "text", idx, texts, violations)
        elif kind == "paragraph":
            # Accept either `runs` (schema) or a flat `text` field (some models prefer it).
            runs = b.get("runs") or []
            for ri, r in enumerate(runs):
                _check_text(r.get("text", ""), f"runs[{ri}].text", idx, texts, violations)
            if not runs and b.get("text"):
                _check_text(b["text"], "text", idx, texts, violations)
        elif kind == "image":
            sid = b.get("manifest_shape_id", "")
            shape = by_id.get(sid)
            if shape is None:
                violations.append(
                    Violation(idx, "manifest_shape_id", sid, "id does not resolve to any manifest shape")
                )
            elif shape.get("kind") not in IMAGE_LIKE_KINDS:
                violations.append(
                    Violation(
                        idx,
                        "manifest_shape_id",
                        sid,
                        f"shape kind {shape.get('kind')!r} is not image-like (Picture/Group/GraphicFrame)",
                    )
                )
            _check_text(b.get("alt", ""), "alt", idx, texts, violations)
            if b.get("caption"):
                _check_text(b["caption"], "caption", idx, texts, violations)
            callout = b.get("callout")
            if callout and callout.get("text"):
                _check_text(callout["text"], "callout.text", idx, texts, violations)
        elif kind == "image_row":
            if b.get("caption"):
                _check_text(b["caption"], "caption", idx, texts, violations)
            for ii, item in enumerate(b.get("images", []) or []):
                sid = item.get("manifest_shape_id", "")
                shape = by_id.get(sid)
                if shape is None:
                    violations.append(
                        Violation(idx, f"images[{ii}].manifest_shape_id", sid, "id does not resolve to any manifest shape")
                    )
                elif shape.get("kind") not in IMAGE_LIKE_KINDS:
                    violations.append(
                        Violation(
                            idx,
                            f"images[{ii}].manifest_shape_id",
                            sid,
                            f"shape kind {shape.get('kind')!r} is not image-like (Picture/Group/GraphicFrame)",
                        )
                    )
                _check_text(item.get("alt", ""), f"images[{ii}].alt", idx, texts, violations)
                if item.get("caption"):
                    _check_text(item["caption"], f"images[{ii}].caption", idx, texts, violations)
                callout = item.get("callout")
                if callout and callout.get("text"):
                    _check_text(callout["text"], f"images[{ii}].callout.text", idx, texts, violations)
        elif kind == "list":
            for ii, item in enumerate(b.get("items", []) or []):
                _check_text(item, f"items[{ii}]", idx, texts, violations)
        elif kind == "table":
            for ri, row in enumerate(b.get("rows", []) or []):
                for ci, cell in enumerate(row):
                    if isinstance(cell, str):
                        _check_text(cell, f"rows[{ri}][{ci}]", idx, texts, violations)
                        continue
                    if not isinstance(cell, dict):
                        violations.append(
                            Violation(
                                idx,
                                f"rows[{ri}][{ci}]",
                                cell,
                                "table cell must be a string or an object",
                            )
                        )
                        continue
                    sid = cell.get("manifest_shape_id")
                    if sid:
                        shape = by_id.get(sid)
                        if shape is None:
                            violations.append(
                                Violation(
                                    idx,
                                    f"rows[{ri}][{ci}].manifest_shape_id",
                                    sid,
                                    "id does not resolve to any manifest shape",
                                )
                            )
                        elif shape.get("kind") not in IMAGE_LIKE_KINDS:
                            violations.append(
                                Violation(
                                    idx,
                                    f"rows[{ri}][{ci}].manifest_shape_id",
                                    sid,
                                    f"shape kind {shape.get('kind')!r} is not image-like (Picture/Group/GraphicFrame)",
                                )
                            )
                        elif not shape.get("image_path"):
                            # Some image-like shapes (e.g. a Group of subshapes with
                            # no rasterized crop) cannot be referenced from a table
                            # image cell because there is nothing to render. The
                            # model should fall back to text or to a different shape.
                            violations.append(
                                Violation(
                                    idx,
                                    f"rows[{ri}][{ci}].manifest_shape_id",
                                    sid,
                                    "shape has no image_path; cannot be used as a table image cell",
                                )
                            )
                    if cell.get("text"):
                        _check_text(cell["text"], f"rows[{ri}][{ci}].text", idx, texts, violations)
                    if cell.get("alt"):
                        _check_text(cell["alt"], f"rows[{ri}][{ci}].alt", idx, texts, violations)
                    if cell.get("caption"):
                        _check_text(cell["caption"], f"rows[{ri}][{ci}].caption", idx, texts, violations)
        elif kind == "quote":
            _check_text(b.get("text", ""), "text", idx, texts, violations)
        elif kind == "divider":
            pass
        else:
            violations.append(Violation(idx, "kind", kind, f"unknown block kind"))

    return violations


# ---- Call -----------------------------------------------------------------


def _png_to_data_url(path: Path) -> str:
    b = Path(path).read_bytes()
    return f"data:image/png;base64,{base64.b64encode(b).decode('ascii')}"


def _slim_manifest(manifest: dict[str, Any]) -> dict[str, Any]:
    """Strip fields the VLM doesn't need: paragraphs, font runs, EMU coords, blob paths.

    Keeps the validator-relevant `text` field intact and the `image_path` so the
    model can reference image-like shapes. bbox_frac stays in normalized form so
    the model can reason about position (e.g. binding callouts to their target)."""
    keep_keys = ("id", "kind", "bbox_frac", "text", "image_path", "is_master_inherited")
    slim_shapes = []
    for s in manifest.get("shapes", []):
        slim_shapes.append({k: s[k] for k in keep_keys if k in s})
    return {
        "slide_index": manifest.get("slide_index"),
        "shapes": slim_shapes,
    }


def _build_user_content(
    slide_png: Path,
    manifest: dict[str, Any],
    notes: str,
    chrome: dict[str, Any],
) -> list[dict[str, Any]]:
    """User message content: image part + text part with structured inputs."""
    text_part = (
        f"manifest:\n{json.dumps(_slim_manifest(manifest), ensure_ascii=False)}\n\n"
        f"notes:\n{notes}\n\n"
        f"chrome:\n{json.dumps(chrome, ensure_ascii=False)}"
    )
    return [
        {
            "type": "image_url",
            "image_url": {"url": _png_to_data_url(slide_png)},
        },
        {"type": "text", "text": text_part},
    ]


def _make_client() -> OpenAI:
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError(
            "OPENAI_API_KEY is not set. Export it (and OPENAI_BASE_URL if not using the "
            "Moonshot default) before running ppt2md."
        )
    base_url = os.environ.get("OPENAI_BASE_URL", DEFAULT_BASE_URL)
    return OpenAI(
        api_key=api_key,
        base_url=base_url,
        timeout=REQUEST_TIMEOUT_S,
        max_retries=SDK_MAX_RETRIES,
    )


def _extract_tool_call(response) -> tuple[str, dict[str, Any]] | None:
    """Return (tool_call_id, arguments_dict) for the first emit_slide_doc call, or None."""
    if not response.choices:
        return None
    msg = response.choices[0].message
    tool_calls = getattr(msg, "tool_calls", None) or []
    for tc in tool_calls:
        if tc.function.name == "emit_slide_doc":
            try:
                args = json.loads(tc.function.arguments or "{}")
            except json.JSONDecodeError:
                args = {}
            return tc.id, args
    return None


def _redact_image_urls(messages: list[dict[str, Any]], slide_png: Path) -> list[dict[str, Any]]:
    """Return a deep copy of `messages` with base64 image data elided for readability."""
    redacted: list[dict[str, Any]] = []
    for m in messages:
        copy: dict[str, Any] = {**m}
        c = copy.get("content")
        if isinstance(c, list):
            new_c: list[dict[str, Any]] = []
            for part in c:
                if isinstance(part, dict) and part.get("type") == "image_url":
                    url = (part.get("image_url") or {}).get("url", "")
                    placeholder = f"<elided base64; original file: {slide_png.name} ({len(url)} chars in data URL)>"
                    new_c.append({"type": "image_url", "image_url": {"url": placeholder}})
                else:
                    new_c.append(part)
            copy["content"] = new_c
        redacted.append(copy)
    return redacted


def _dump_prompt(
    debug_dir: Path,
    slide_index: int,
    messages: list[dict[str, Any]],
    slide_png: Path,
    model: str,
    n_attempts: int,
) -> None:
    out = {
        "model": model,
        "temperature": 1,
        "tool_choice": "auto",
        "n_attempts": n_attempts,
        "tools": [EMIT_SLIDE_DOC_TOOL],
        "messages": _redact_image_urls(messages, slide_png),
    }
    (debug_dir / f"prompt-{slide_index}.json").write_text(
        json.dumps(out, indent=2, ensure_ascii=False), encoding="utf-8"
    )


def _usage_dict(response) -> dict[str, Any]:
    """Pull token usage off the response (Moonshot follows the OpenAI shape)."""
    u = getattr(response, "usage", None)
    if u is None:
        return {}
    out: dict[str, Any] = {}
    for k in ("prompt_tokens", "completion_tokens", "total_tokens"):
        v = getattr(u, k, None)
        if v is not None:
            out[k] = v
    # Reasoning/thinking tokens vary by provider; capture if present.
    details = getattr(u, "completion_tokens_details", None)
    if details is not None:
        rt = getattr(details, "reasoning_tokens", None)
        if rt is not None:
            out["reasoning_tokens"] = rt
    return out


def call_vlm(
    slide_png: Path,
    manifest: dict[str, Any],
    notes: str,
    chrome: dict[str, Any],
    model: str | None = None,
    debug_dir: Path | None = None,
    slide_index: int | None = None,
    logger: Any | None = None,
) -> tuple[dict[str, Any], int]:
    """Call the VLM, validate, retry up to MAX_RETRIES, return (slide_doc, n_attempts).

    If `debug_dir` and `slide_index` are provided, the final messages array
    (including any retry turns) is written to `<debug_dir>/prompt-{slide_index}.json`
    with image base64 data elided. Useful for offline inspection of what was sent.

    If `logger` is provided (RunLogger), per-attempt events are emitted with
    wall time, model, and token usage.

    Raises ValidationFailure on persistent validation failure.
    """
    import time

    client = _make_client()
    model = model or os.environ.get("PPT2MD_MODEL", DEFAULT_MODEL)

    user_content = _build_user_content(slide_png, manifest, notes, chrome)

    messages: list[dict[str, Any]] = [
        {"role": "system", "content": SYSTEM_PROMPT},
        {"role": "user", "content": user_content},
    ]

    last_violations: list[Violation] = []
    n_attempts = 0
    cumulative_usage = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0, "reasoning_tokens": 0}
    cumulative_seconds = 0.0
    for attempt in range(MAX_RETRIES + 1):
        n_attempts += 1
        t0 = time.monotonic()
        try:
            response = client.chat.completions.create(
                model=model,
                messages=messages,
                tools=[EMIT_SLIDE_DOC_TOOL],
                tool_choice="auto",
                temperature=1,
            )
        except (APITimeoutError, APIConnectionError) as exc:
            elapsed = time.monotonic() - t0
            cumulative_seconds += elapsed
            err_kind = type(exc).__name__
            if logger is not None:
                logger.event(
                    "vlm.network_error",
                    f"      slide {slide_index} attempt {n_attempts}: {err_kind} after "
                    f"{elapsed:.1f}s — {exc}",
                    slide_index=slide_index,
                    attempt=n_attempts,
                    model=model,
                    seconds=round(elapsed, 3),
                    error=err_kind,
                )
            last_violations = [
                Violation(
                    block_index=-1,
                    field="network",
                    value=err_kind,
                    reason=f"{err_kind}: {exc}",
                )
            ]
            continue
        elapsed = time.monotonic() - t0
        cumulative_seconds += elapsed
        usage = _usage_dict(response)
        for k, v in usage.items():
            cumulative_usage[k] = cumulative_usage.get(k, 0) + v
        if logger is not None:
            logger.event(
                "vlm.attempt",
                f"      slide {slide_index} attempt {n_attempts}: {elapsed:.1f}s, "
                f"in={usage.get('prompt_tokens', '?')} out={usage.get('completion_tokens', '?')}"
                + (f" thinking={usage['reasoning_tokens']}" if "reasoning_tokens" in usage else ""),
                slide_index=slide_index,
                attempt=n_attempts,
                model=model,
                seconds=round(elapsed, 3),
                usage=usage,
            )
        extracted = _extract_tool_call(response)
        if extracted is None:
            # Model failed to call the tool. Log this as a single violation and retry.
            last_violations = [
                Violation(
                    block_index=-1,
                    field="tool_call",
                    value=None,
                    reason="model did not call emit_slide_doc",
                )
            ]
        else:
            tool_call_id, slide_doc = extracted
            last_violations = validate(slide_doc, manifest)
            if not last_violations:
                if debug_dir is not None and slide_index is not None:
                    _dump_prompt(debug_dir, slide_index, messages, slide_png, model, n_attempts)
                if logger is not None:
                    logger.event(
                        "vlm.ok",
                        f"      slide {slide_index} OK after {n_attempts} attempt(s) "
                        f"({cumulative_seconds:.1f}s total, "
                        f"in={cumulative_usage['prompt_tokens']} out={cumulative_usage['completion_tokens']})",
                        slide_index=slide_index,
                        attempts=n_attempts,
                        seconds_total=round(cumulative_seconds, 3),
                        usage_total=cumulative_usage,
                    )
                return slide_doc, n_attempts
            if logger is not None:
                # Include the offending value (truncated) so the user can compare
                # it to the manifest text without opening prompt-{i}.json.
                def _short(v: Any, n: int = 80) -> str:
                    if v is None:
                        return "None"
                    s = str(v).replace("\n", "⏎").replace("\r", "")
                    return s if len(s) <= n else s[:n] + "…"

                preview = "; ".join(
                    f"[{v.block_index}].{v.field}={_short(v.value)!r}"
                    for v in last_violations[:3]
                )
                more = f" (+{len(last_violations) - 3} more)" if len(last_violations) > 3 else ""
                logger.event(
                    "vlm.invalid",
                    f"      slide {slide_index} attempt {n_attempts} invalid "
                    f"({len(last_violations)} violation(s)): {preview}{more}",
                    slide_index=slide_index,
                    attempt=n_attempts,
                    n_violations=len(last_violations),
                    violations=[v.to_dict() for v in last_violations],
                )
            # Append the assistant turn and a tool-result message echoing the violations.
            messages.append(
                {
                    "role": "assistant",
                    "tool_calls": [
                        {
                            "id": tool_call_id,
                            "type": "function",
                            "function": {
                                "name": "emit_slide_doc",
                                "arguments": json.dumps(slide_doc),
                            },
                        }
                    ],
                    "content": "",
                }
            )
            messages.append(
                {
                    "role": "tool",
                    "tool_call_id": tool_call_id,
                    "content": json.dumps(
                        {"violations": [v.to_dict() for v in last_violations]},
                        ensure_ascii=False,
                    ),
                }
            )
            continue

        # tool_call missing — append a user-style nudge and retry.
        messages.append(
            {
                "role": "user",
                "content": "You must call the emit_slide_doc function exactly once. Try again.",
            }
        )

    if debug_dir is not None and slide_index is not None:
        _dump_prompt(debug_dir, slide_index, messages, slide_png, model, n_attempts)
    if logger is not None:
        logger.event(
            "vlm.failed",
            f"      slide {slide_index} FAILED after {n_attempts} attempt(s)",
            slide_index=slide_index,
            attempts=n_attempts,
            seconds_total=round(cumulative_seconds, 3),
            usage_total=cumulative_usage,
            n_violations=len(last_violations),
        )
    raise ValidationFailure(last_violations)
