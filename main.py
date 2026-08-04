"""启动入口：无参数启动 GUI；带参数时走命令行。"""

from __future__ import annotations

import sys


def main() -> None:
    if len(sys.argv) > 1:
        from cli import run_cli

        raise SystemExit(run_cli())
    from app import main as gui_main

    gui_main()


if __name__ == "__main__":
    main()
