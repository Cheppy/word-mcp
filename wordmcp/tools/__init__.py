"""Tool groups. Each module exposes register(mcp, service) and adds nothing global."""

from . import escape, measure, modify, session

REGISTRARS = (session.register, measure.register, modify.register)

__all__ = ["REGISTRARS", "escape", "measure", "modify", "session"]
