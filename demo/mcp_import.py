"""Read-only example: python demo/mcp_import.py [--window TITLE | --list-tools]."""

import argparse

import easy_uiauto.mcp as tools
from easy_uiauto.mcp import get_control_tree, list_windows


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--window", help="Exact title returned by list_windows")
    group.add_argument("--list-tools", action="store_true")
    args = parser.parse_args()
    if args.list_tools:
        print("\n".join(tools.__all__))
    elif args.window:
        print(get_control_tree(args.window, max_nodes=500, timeout_seconds=15))
    else:
        print(list_windows())


if __name__ == "__main__":
    main()
