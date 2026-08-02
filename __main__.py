"""Allows `python -m excelmcp` to start the MCP server.

Agents launched from a GUI often do not inherit the shell PATH, so the console
script may be unreachable even when the package is installed. This module gives
`agents.resolve_command()` a fallback that always works.
"""

from excelmcp.main import main

if __name__ == "__main__":
    main()
