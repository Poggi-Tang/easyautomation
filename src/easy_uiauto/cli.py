from __future__ import annotations

import argparse

from . import __version__


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="easy-uiauto",
        description="Windows UI automation recorder and action runner",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    subparsers = parser.add_subparsers(dest="command")
    record_parser = subparsers.add_parser("record", help="record mouse and keyboard actions")
    record_parser.add_argument(
        "--no-write",
        action="store_true",
        help="record actions without generating a Python file",
    )
    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.command == "record":
        from .record import run_record

        run_record(write_file=not args.no_write)
        return 0
    parser.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
