"""Engine selection. Dependencies are passed in, never read from module scope."""

from __future__ import annotations

import os
import sys
from collections.abc import Callable, Mapping

from ..engine import DocumentEngine

BACKEND_VARIABLE = "WORD_MCP_BACKEND"
PORT_VARIABLE = "WORD_MCP_UNO_PORT"

PLATFORM_DEFAULT = {"darwin": "applescript", "win32": "com"}
ALIASES = {
    "mac": "applescript",
    "macos": "applescript",
    "applescript": "applescript",
    "win": "com",
    "windows": "com",
    "com": "com",
    "linux": "uno",
    "uno": "uno",
    "libreoffice": "uno",
}


def _applescript(env: Mapping[str, str]) -> DocumentEngine:
    from .applescript import AppleScriptEngine

    return AppleScriptEngine()


def _com(env: Mapping[str, str]) -> DocumentEngine:
    from .com import ComEngine

    return ComEngine()


def _uno(env: Mapping[str, str]) -> DocumentEngine:
    from .uno import UnoEngine

    return UnoEngine(port=int(env.get(PORT_VARIABLE, "2002")))


BUILDERS: dict[str, Callable[[Mapping[str, str]], DocumentEngine]] = {
    "applescript": _applescript,
    "com": _com,
    "uno": _uno,
}


def resolve_name(
    requested: str | None = None,
    *,
    platform: str = sys.platform,
    env: Mapping[str, str] | None = None,
) -> str:
    """Name of the engine to use: an explicit request, the environment, or the platform."""
    environment = os.environ if env is None else env
    choice = (requested or environment.get(BACKEND_VARIABLE) or "").strip().lower()
    if not choice:
        return PLATFORM_DEFAULT.get(platform, "uno")
    if choice not in ALIASES:
        raise ValueError(f"unknown engine {choice!r}; expected one of {sorted(set(ALIASES))}")
    return ALIASES[choice]


def create_engine(
    requested: str | None = None,
    *,
    platform: str = sys.platform,
    env: Mapping[str, str] | None = None,
) -> DocumentEngine:
    """Build the engine for this machine without touching global state."""
    environment = os.environ if env is None else env
    return BUILDERS[resolve_name(requested, platform=platform, env=environment)](environment)
