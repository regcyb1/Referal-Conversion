"""Entry point for the packaged desktop app (.exe / .app).

Double-click opens the window. Given arguments, it behaves like `convert.py`
on the command line (e.g. for scripted use or build checks).
"""

import os
import sys

# A windowed .exe has no console: give print() somewhere harmless to write.
for _name in ("stdout", "stderr"):
    if getattr(sys, _name) is None:
        setattr(sys, _name, open(os.devnull, "w"))


def main():
    # macOS may pass a "-psn_..." process id when launched from Finder.
    args = [arg for arg in sys.argv[1:] if not arg.startswith("-psn")]
    if args:
        from convert import main as cli
        return cli(args)
    from desktop import launch
    launch()
    return 0


if __name__ == "__main__":
    sys.exit(main())
