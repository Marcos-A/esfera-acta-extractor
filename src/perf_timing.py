"""
Lightweight timing helpers for measuring conversion stages.
"""

from __future__ import annotations

import json
import os
import time
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Iterator


def perf_timing_enabled() -> bool:
    """Enable stage timing only when explicitly requested via environment."""
    return os.getenv("PERF_TIMING_ENABLED", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


@dataclass
class TimingRecorder:
    """Collect stage timings and emit a single structured log line when enabled."""

    label: str
    enabled: bool = field(default_factory=perf_timing_enabled)
    _started_at: float = field(default_factory=time.perf_counter)
    _stages: list[tuple[str, float]] = field(default_factory=list)

    @contextmanager
    def measure(self, stage_name: str) -> Iterator[None]:
        """Measure the duration of one named stage."""
        stage_started_at = time.perf_counter()
        try:
            yield
        finally:
            if self.enabled:
                self._stages.append((stage_name, time.perf_counter() - stage_started_at))

    def log(self, **extra: object) -> None:
        """Emit one JSON log line that is easy to compare across benchmark runs."""
        if not self.enabled:
            return

        # One structured line per operation is easier to diff and grep than many small
        # print statements scattered across the conversion pipeline.
        payload: dict[str, object] = {
            "event": "perf_timing",
            "label": self.label,
            "total_seconds": round(time.perf_counter() - self._started_at, 6),
            "stages": [
                {"name": stage_name, "seconds": round(duration, 6)}
                for stage_name, duration in self._stages
            ],
        }
        if extra:
            payload["context"] = extra

        print("[PERF] " + json.dumps(payload, ensure_ascii=True, sort_keys=True))
