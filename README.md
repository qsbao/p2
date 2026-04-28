# ppt2md

Generic PowerPoint → Markdown pipeline. Renders each slide, extracts a
structural manifest of every shape, sends both to a vision LLM (Kimi by
default), validates the response against the manifest, and writes a markdown
file plus the cropped images it links to.

## Install

External binaries — resolved via `$PATH` automatically:

| binary | macOS | Linux (Debian/Ubuntu) |
|---|---|---|
| LibreOffice (`soffice`) | `brew install --cask libreoffice` | `sudo apt install libreoffice` |
| Poppler (`pdftoppm`) | `brew install poppler` | `sudo apt install poppler-utils` |

Override the resolved paths with `PPT2MD_SOFFICE` / `PPT2MD_PDFTOPPM` if you
need a non-default install.

Python ≥ 3.10 with three packages:

```sh
pip install openai python-pptx pillow
```

Clone the repo somewhere on your `PYTHONPATH` (or run from its parent dir):

```sh
git clone <repo> ppt2md
cd ppt2md/..   # so `python -m ppt2md` resolves
```

Set the LLM credentials. Defaults target Moonshot Kimi; any OpenAI-compatible
vision + tool-use endpoint works:

```sh
export OPENAI_API_KEY=sk-...
export OPENAI_BASE_URL=https://api.moonshot.cn/v1   # optional; this is the default
export PPT2MD_MODEL=kimi-k2.6                       # optional; this is the default
```

## Usage

```sh
python -m ppt2md [--stream] [--upload {none,s3,cmd}] <pptx> <out_dir>
```

Example:

```sh
python -m ppt2md deck.pptx /tmp/run-1
```

### Tunable constants (env vars)

Pipeline thresholds and timeouts read from the environment at import time;
unset means use the default. Invalid values fail loud rather than silently
falling back, so a typo is caught early.

| Env var | Default | Used in | Meaning |
|---|---|---|---|
| `PPT2MD_GROUP_AREA_THRESHOLD` | `0.05` | `extract.py` | Min fraction of slide area for a Group to count as image-like |
| `PPT2MD_CROP_PAD_FRAC` | `0.02` | `extract.py` | Padding around each crop, as a fraction of slide width |
| `PPT2MD_MAX_RETRIES` | `4` | `vlm.py` | Outer validator-driven retry budget per slide |
| `PPT2MD_REQUEST_TIMEOUT_S` | `120.0` | `vlm.py` | Per-request HTTP timeout (seconds) |
| `PPT2MD_VLM_CONCURRENCY` | `8` | `cli.py` | Max in-flight slide VLM calls |

### Streaming mode

`--stream` emits NDJSON events on stdout (one per line) so downstream tools
can start consuming markdown the moment each slide is ready. Slides are
emitted in slide-index order via a reorder buffer (the VLM still runs in
parallel underneath). The full `<stem>.md` and all debug artifacts are still
written to disk — the stream is purely additive.

```sh
python -m ppt2md --stream deck.pptx /tmp/run-1 \
    | python ../scripts/stream_consumer.py
```

Event schema:

```jsonc
{"type":"start", "stem":"deck", "n_slides":42, "out_dir":"/tmp/run-1",
 "stem_dir":"deck", "media_dir":"deck/media"}

{"type":"slide", "slide_index":1, "markdown":"# ...\n",
 "media":["deck/media/slide1-full.png", "deck/media/slide1-fig1.png", ...]}

// ... one slide event per slide, in slide-index order ...

{"type":"done", "md_path":"/tmp/run-1/deck.md", "seconds_total":312.4,
 "chrome_audit_path":"/tmp/run-1/deck.debug/chrome_dropped.md"}
```

On validation failure after retries the stream terminates with:

```jsonc
{"type":"error", "stage":"vlm", "message":"validation failed after retries",
 "violations_path":"/tmp/run-1/deck.debug/validation_errors.json",
 "n_violations":3}
```

Stderr keeps the human/debug log; stdout is reserved for NDJSON when
`--stream` is on. Media paths in `slide` events are relative to `out_dir`
from the `start` event, matching the image links inside the markdown.

### Remote image hosting (`--upload`)

By default the rendered markdown references local files under
`<out_dir>/<stem>/media/`. For deployments that need the markdown to be
self-contained over HTTP — e.g. piped into a chat thread, indexed by an
LLM service, embedded in a CMS — `--upload` mirrors every figure (cropped
shapes + the full-slide reference image) to remote storage and rewrites the
markdown image links to the returned URLs. Local files are still written so
re-runs and debug remain straightforward.

Built-in providers:

| `--upload` | Required env vars | URL returned |
|---|---|---|
| `none` (default) | — | `<stem>/media/<file>` (local relative) |
| `s3` | `PPT2MD_S3_BUCKET` (+ optional `PPT2MD_S3_PREFIX`, `PPT2MD_S3_PUBLIC_BASE`, `PPT2MD_S3_ACL`, `PPT2MD_S3_CONTENT_TYPE`) | `https://<base>/…` if `PPT2MD_S3_PUBLIC_BASE` is set, else `s3://<bucket>/<prefix>/<key>` |
| `cmd` | `PPT2MD_UPLOAD_CMD` (shell template with `{src}` / `{key}`; last stdout line = URL) | whatever the command echoes |

S3 example:

```sh
pip install boto3
export PPT2MD_S3_BUCKET=my-decks
export PPT2MD_S3_PREFIX=2026/q1
export PPT2MD_S3_PUBLIC_BASE=https://cdn.example.com
python -m ppt2md --upload s3 deck.pptx /tmp/run-1
```

Custom-command example (any HTTP store):

```sh
export PPT2MD_UPLOAD_CMD='curl -s -F file=@{src} -F key={key} https://my-uploader/internal | jq -r .url'
python -m ppt2md --upload cmd deck.pptx /tmp/run-1
```

Object keys mirror the local layout (`<stem>/media/<file>`), so re-runs to a
fresh `<out_dir>` upload to deterministic URLs. To plug in a custom Python
uploader, implement the `Uploader` protocol in `ppt2md/upload.py` (one method:
`upload(src: Path, key: str) -> str`) and call `make_uploader` directly.

For input `deck.pptx` and `<out_dir> = /tmp/run-1`, the pipeline writes:

```
/tmp/run-1/
  deck.md                         # final markdown
  deck/media/                     # cropped figures referenced from deck.md
    slide1-fig1.png
    ...
  deck.debug/                     # inspection material (safe to delete)
    slide-1.png                   # 150-DPI rendered slide
    manifest-1.json               # structural extract of every shape
    notes-1.txt                   # speaker notes
    chrome-1.json                 # chrome flags
    prompt-1.json                 # messages sent to the LLM
    slide_doc-1.json              # LLM tool-call output
    chrome_dropped.md             # audit: shapes dropped + why
    validation_errors.json        # only on validation failure
```

Wall time is typically 30–90 s per slide (Kimi is a thinking model). Up to 2
retries if the validator catches hallucinated text or invalid image refs.

Exit codes: `0` success, `1` pipeline error (see stderr), `2` argument error.
Re-runs overwrite outputs; there is no caching.
