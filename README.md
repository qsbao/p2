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
python -m ppt2md <pptx> <out_dir>
```

Example:

```sh
python -m ppt2md deck.pptx /tmp/run-1
```

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
