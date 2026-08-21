"""Structured logging utilities for PowerPoint MCP server."""

import logging
import os
import sys
from typing import Optional


def get_logger(name: str = "powerpoint_mcp") -> logging.Logger:
    """Get or configure a logger with consistent formatting and stream handler.

    Args:
        name: Logger module or component name.

    Returns:
        Configured logging.Logger instance.
    """
    logger = logging.getLogger(name)

    # If handlers already configured, return immediately to avoid duplication
    if logger.handlers:
        return logger

    # Determine log level from environment variable
    level_str = os.environ.get("PPT_LOG_LEVEL", "INFO").upper()
    level = getattr(logging, level_str, logging.INFO)
    logger.setLevel(level)

    # Standard stream handler sending logs to stderr (safe for MCP stdio protocol)
    handler = logging.StreamHandler(sys.stderr)
    handler.setLevel(level)

    formatter = logging.Formatter(
        fmt="%(asctime)s [%(levelname)s] [%(name)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    # Prevent propagating to root logger to avoid double logging
    logger.propagate = False

    return logger
