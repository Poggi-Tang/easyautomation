"""Command-line interface for scanned application UI knowledge."""

from __future__ import annotations

import argparse
import json
import sys

from easy_uiauto import __version__

from . import configuration, knowledge, scanner, ui_cli


def _vision_settings(args) -> tuple[str, str, str]:
    api_url = (
        getattr(args, "vision_url", "")
        or configuration._existing_vision_value(configuration.VISION_API_URL)
    ).strip()
    api_key = configuration._existing_vision_value(configuration.VISION_API_KEY)
    model = (
        getattr(args, "vision_model", "")
        or configuration._existing_vision_value(configuration.VISION_MODEL)
    ).strip()
    if not api_url or not api_key or not model:
        raise RuntimeError(
            "Scanning requires EASY_UIAUTO_VISION_API_URL, "
            "EASY_UIAUTO_VISION_API_KEY, and EASY_UIAUTO_VISION_MODEL"
        )
    return api_url, api_key, model


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="easy_uiauto_ui",
        description=(
            "Learn real control meanings from page context and run verified semantic UI commands."
        ),
    )
    parser.add_argument("--version", action="version", version=f"easy_uiauto_ui {__version__}")
    commands = parser.add_subparsers(dest="subcommand", required=True)

    scan = commands.add_parser("scan", help="Scan a visible application window.")
    scan.add_argument("window_name")
    scan.add_argument("--max-depth", type=int, default=12)
    scan.add_argument("--max-controls", type=int, default=3000)
    scan.add_argument("--verify-limit", type=int, default=500)
    scan.add_argument("--vision-url", default="")
    scan.add_argument("--vision-model", default="")

    commands.add_parser("apps", help="List scanned applications.")

    search = commands.add_parser("search", help="Search an application's controls.")
    search.add_argument("app_id")
    search.add_argument("query", nargs="?", default="")
    search.add_argument("--include-quarantine", action="store_true")
    search.add_argument("--limit", type=int, default=50)

    list_commands = commands.add_parser("commands", help="List verified UI commands.")
    list_commands.add_argument("app_id")
    list_commands.add_argument("--page", default="")

    run = commands.add_parser("run", help="Run one verified UI command.")
    run.add_argument("app_id")
    run.add_argument("command")
    run.add_argument("--text", default="")
    run.add_argument("--confirm", action="store_true")

    teach = commands.add_parser("teach", help="Teach or correct a control's real function.")
    teach.add_argument("app_id")
    teach.add_argument("control_id")
    teach.add_argument("semantic_name")
    teach.add_argument("intent")
    teach.add_argument("description")
    teach.add_argument("--actions", default="")
    teach.add_argument("--aliases", default="")
    teach.add_argument("--risk", default="safe")
    teach.add_argument("--requires-confirmation", action="store_true")

    rebuild = commands.add_parser("reindex", help="Rebuild cache and catalog from Markdown.")
    rebuild.add_argument("app_id")
    return parser


def main(argv: list[str] | None = None) -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        if args.subcommand == "scan":
            api_url, api_key, model = _vision_settings(args)
            result = scanner.scan_window(
                args.window_name,
                api_url,
                api_key,
                model,
                __version__,
                max_depth=args.max_depth,
                max_controls=args.max_controls,
                verify_limit=args.verify_limit,
                progress=lambda message: print(f"[scan] {message}", flush=True),
            )
        elif args.subcommand == "apps":
            result = knowledge.list_apps()
        elif args.subcommand == "search":
            statuses = None
            if args.include_quarantine:
                statuses = {"verified", "observed", "suspect", "quarantined"}
            result = knowledge.search_controls(
                knowledge.app_dir(args.app_id),
                args.query,
                statuses,
                args.limit,
            )
        elif args.subcommand == "commands":
            result = knowledge.available_commands(
                knowledge.app_dir(args.app_id),
                args.page,
            )
        elif args.subcommand == "run":
            result = ui_cli.execute(
                knowledge.app_dir(args.app_id),
                args.command,
                args.text,
                args.confirm,
            )
        elif args.subcommand == "teach":
            selected_actions = [
                value.strip() for value in args.actions.split(",") if value.strip()
            ]
            result = knowledge.teach_control(
                knowledge.app_dir(args.app_id),
                args.control_id,
                args.semantic_name,
                args.intent,
                args.description,
                selected_actions or None,
                [value.strip() for value in args.aliases.split(",") if value.strip()],
                args.risk,
                args.requires_confirmation,
            )
        else:
            directory = knowledge.app_dir(args.app_id)
            index = knowledge.rebuild_index(directory)
            catalog = knowledge.write_command_catalog(directory)
            result = {
                "ok": True,
                "controls": len(index["controls"]),
                "commands": len(knowledge.available_commands(directory)),
                "catalog": str(catalog),
            }
    except (KeyError, RuntimeError, ValueError) as error:
        parser.error(str(error))
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
