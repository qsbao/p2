"""Image-uploader plugins.

The pipeline writes cropped figures to `<out_dir>/<stem>/media/`. When a
caller wants the rendered markdown to reference remote URLs (S3, a CDN, an
internal asset service…) instead of those local paths, an Uploader is given
each local file and returns the URL to embed.

Built-ins:
    - NoopUploader: returns the key as-is (default; markdown stays local-relative)
    - S3Uploader:   uploads via boto3; URL = PPT2MD_S3_PUBLIC_BASE/{key}, or
                    s3://{bucket}/{key} when no public base is configured
    - CmdUploader:  runs PPT2MD_UPLOAD_CMD shell template; URL = last stdout line

Custom uploaders only need to implement `upload(src, key) -> url`.
"""

from __future__ import annotations

import os
import shlex
import subprocess
from pathlib import Path
from typing import Protocol


class Uploader(Protocol):
    def upload(self, src: Path, key: str) -> str:
        """Upload `src` (a local file) under the logical `key` and return the
        URL/path that should be embedded in the rendered markdown."""
        ...


class NoopUploader:
    """No-op: returns the key (a relative path) unchanged. Use when the
    markdown should keep referencing local files."""

    def upload(self, src: Path, key: str) -> str:
        return key


class S3Uploader:
    """Upload to S3 via boto3. Configure via environment:

        PPT2MD_S3_BUCKET       (required) target bucket name
        PPT2MD_S3_PREFIX       (optional) key prefix, e.g. "decks/2026"
        PPT2MD_S3_PUBLIC_BASE  (optional) public URL base, e.g.
                               "https://cdn.example.com" — when set, the
                               returned URL is "{base}/{prefix}/{key}";
                               otherwise it is "s3://{bucket}/{prefix}/{key}"
        PPT2MD_S3_ACL          (optional) canned ACL, e.g. "public-read"
        PPT2MD_S3_CONTENT_TYPE (optional) override; default "image/png"

    Standard AWS credentials (`AWS_ACCESS_KEY_ID`, `AWS_PROFILE`, IAM role,
    etc.) are picked up by boto3 in the usual way.
    """

    def __init__(self) -> None:
        bucket = os.environ.get("PPT2MD_S3_BUCKET")
        if not bucket:
            raise RuntimeError("S3 uploader requires PPT2MD_S3_BUCKET")
        try:
            import boto3  # type: ignore
        except ImportError as e:
            raise RuntimeError(
                "S3 uploader requires the `boto3` package; run `pip install boto3`"
            ) from e
        self.bucket = bucket
        self.prefix = (os.environ.get("PPT2MD_S3_PREFIX") or "").strip("/")
        self.public_base = (os.environ.get("PPT2MD_S3_PUBLIC_BASE") or "").rstrip("/")
        self.acl = os.environ.get("PPT2MD_S3_ACL") or None
        self.content_type = os.environ.get("PPT2MD_S3_CONTENT_TYPE") or "image/png"
        self._client = boto3.client("s3")

    def _full_key(self, key: str) -> str:
        key = key.lstrip("/")
        return f"{self.prefix}/{key}" if self.prefix else key

    def upload(self, src: Path, key: str) -> str:
        full_key = self._full_key(key)
        extra: dict[str, str] = {"ContentType": self.content_type}
        if self.acl:
            extra["ACL"] = self.acl
        self._client.upload_file(str(src), self.bucket, full_key, ExtraArgs=extra)
        if self.public_base:
            return f"{self.public_base}/{full_key}"
        return f"s3://{self.bucket}/{full_key}"


class CmdUploader:
    """Run a shell template per file. The template comes from
    `PPT2MD_UPLOAD_CMD` and may contain `{src}` (local file path) and `{key}`
    (logical key) placeholders. The URL is the last non-empty line of stdout.

    Example:

        export PPT2MD_UPLOAD_CMD='aws s3 cp {src} s3://my-bucket/{key} > /dev/null \\
            && echo https://cdn.example.com/{key}'

    The placeholders are substituted *before* shell parsing, so paths must not
    contain shell metacharacters. Crops produced by ppt2md never do.
    """

    def __init__(self) -> None:
        tmpl = os.environ.get("PPT2MD_UPLOAD_CMD")
        if not tmpl:
            raise RuntimeError("cmd uploader requires PPT2MD_UPLOAD_CMD")
        self.template = tmpl

    def upload(self, src: Path, key: str) -> str:
        cmd = self.template.format(src=shlex.quote(str(src)), key=shlex.quote(key))
        proc = subprocess.run(
            cmd, shell=True, capture_output=True, text=True, check=False
        )
        if proc.returncode != 0:
            raise RuntimeError(
                f"upload command failed (exit {proc.returncode}) for {src}:\n"
                f"  cmd: {cmd}\n  stderr: {proc.stderr.strip()}"
            )
        lines = [ln.strip() for ln in proc.stdout.splitlines() if ln.strip()]
        if not lines:
            raise RuntimeError(
                f"upload command produced no URL on stdout for {src}\n  cmd: {cmd}"
            )
        return lines[-1]


def make_uploader(name: str) -> Uploader:
    """Instantiate the named uploader. Raises on unknown names or misconfig."""
    name = (name or "none").lower()
    if name in ("none", "noop", ""):
        return NoopUploader()
    if name == "s3":
        return S3Uploader()
    if name == "cmd":
        return CmdUploader()
    raise ValueError(f"unknown uploader: {name!r} (expected: none, s3, cmd)")
