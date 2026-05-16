from __future__ import annotations

import logging


class NoRawNoteFilter(logging.Filter):
    """Prevents accidental high-volume note-like payloads from being emitted."""

    def filter(self, record: logging.LogRecord) -> bool:
        message = record.getMessage()
        return len(message) < 500


def configure_private_logging() -> None:
    logging.getLogger("uvicorn.access").disabled = True
    for logger_name in ("uvicorn", "uvicorn.error", "engine"):
        logger = logging.getLogger(logger_name)
        logger.addFilter(NoRawNoteFilter())

