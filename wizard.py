"""Backwards-compatible shim.

The setup experience now lives in `excelmcp.cli`, which adds agent auto-detection,
subcommands, and non-interactive flags. This module is kept so that anything
pinned to `excelmcp.wizard:main` (an older console-script entry point, a script,
or a doc snippet) keeps working.
"""

from excelmcp.cli import main, run_setup  # noqa: F401

__all__ = ["main", "run_setup"]

if __name__ == "__main__":
    main()
