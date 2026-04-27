"""Tiny structured logger that writes to both stderr (human view) and a JSONL
file in the debug directory (machine view).

Each record is a single JSON object on its own line with a ts (UTC ISO-8601),
event name, and arbitrary fields.
"""

from __future__ import annotations

import json
import sys
import threading
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


class RunLogger:
    def __init__(self, log_path: Path | None) -> None:
        self.log_path = log_path
        self._lock = threading.Lock()
        if log_path is not None:
            log_path.parent.mkdir(parents=True, exist_ok=True)
            log_path.write_text("")  # truncate

    def event(self, event: str, message: str | None = None, **fields: Any) -> None:
        rec = {
            "ts": datetime.now(timezone.utc).isoformat(timespec="milliseconds"),
            "event": event,
        }
        if message is not None:
            rec["message"] = message
        rec.update(fields)
        with self._lock:
            if self.log_path is not None:
                with self.log_path.open("a") as f:
                    f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            if message:
                print(message, file=sys.stderr, flush=True)
