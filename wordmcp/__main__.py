"""Composition root: choose an engine for this machine and serve over stdio."""

from __future__ import annotations

import os
from collections.abc import Mapping

from .engines import create_engine
from .server import create_server
from .service import DocumentService

RAW_VARIABLE = "WORD_MCP_ALLOW_RAW"


def build_service(env: Mapping[str, str] | None = None) -> DocumentService:
    """Construct the service for this machine from the environment."""
    environment = os.environ if env is None else env
    return DocumentService(create_engine(env=environment))


def main() -> None:
    """Wire the default service and run the server."""
    environment = os.environ
    create_server(
        build_service(environment),
        allow_raw=environment.get(RAW_VARIABLE) == "1",
    ).run()


if __name__ == "__main__":
    main()
