"""Bible2PPT entry point.

Launches the Tkinter GUI. All application logic is in :mod:`core`; the UI is in
:mod:`ui.app`. Kept intentionally thin so the same core can be driven by other
front-ends (CLI, tests) later.
"""
from __future__ import annotations

from ui.app import run


def main() -> None:
    run()


if __name__ == "__main__":
    main()
