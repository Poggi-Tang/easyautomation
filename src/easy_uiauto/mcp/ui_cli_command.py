"""Command-line interface for scanned application UI knowledge."""

from __future__ import annotations

import argparse
import json
import sys

from easy_uiauto import __version__

from . import configuration, interaction_learning, knowledge, scanner, ui_cli


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
    scan.add_argument(
        "--strategy",
        choices=("visual-first", "full-uia"),
        default="visual-first",
    )
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
    run.add_argument("--allow-vision-fallback", action="store_true")

    batch = commands.add_parser(
        "batch",
        help="Run a preflighted same-page sequence in one process.",
    )
    batch.add_argument("app_id")
    batch.add_argument(
        "steps_json",
        help=(
            "JSON array of command strings or {command,text} objects; "
            "use @path to read JSON from a file"
        ),
    )
    batch.add_argument("--confirm", action="store_true")
    batch.add_argument("--allow-vision-fallback", action="store_true")

    effect = commands.add_parser(
        "learn-effect",
        help="Execute one command and learn its before/after response.",
    )
    effect.add_argument("app_id")
    effect.add_argument("command")
    effect.add_argument("--text", default="")
    effect.add_argument("--confirm", action="store_true")
    effect.add_argument("--recover", action="store_true")
    effect.add_argument("--maximum-wait", type=float, default=3.0)
    effect.add_argument("--vision-url", default="")
    effect.add_argument("--vision-model", default="")

    explore = commands.add_parser(
        "explore",
        help="Learn direct responses of known reversible commands.",
    )
    explore.add_argument("app_id")
    explore.add_argument("--policy", choices=("safe", "supervised"), default="safe")
    explore.add_argument("--max-actions", type=int, default=10)
    explore.add_argument("--max-depth", type=int, default=3)
    explore.add_argument("--confirm", action="store_true")
    explore.add_argument("--vision-url", default="")
    explore.add_argument("--vision-model", default="")

    interactions = commands.add_parser(
        "interactions",
        help="List learned operation effects.",
    )
    interactions.add_argument("app_id")
    interactions.add_argument("--command", default="")
    interactions.add_argument("--limit", type=int, default=50)

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
                strategy=args.strategy,
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
                args.allow_vision_fallback,
            )
        elif args.subcommand == "batch":
            value = args.steps_json
            if value.startswith("@"):
                with open(value[1:], encoding="utf-8") as stream:
                    value = stream.read()
            steps = json.loads(value)
            batch_args = (knowledge.app_dir(args.app_id), steps, args.confirm)
            result = (
                ui_cli.execute_many(*batch_args, allow_vision_fallback=True)
                if args.allow_vision_fallback
                else ui_cli.execute_many(*batch_args)
            )
        elif args.subcommand == "learn-effect":
            api_url, api_key, model = _vision_settings(args)
            result = interaction_learning.learn_command_effect(
                knowledge.app_dir(args.app_id),
                args.command,
                api_url,
                api_key,
                model,
                __version__,
                text=args.text,
                confirm=args.confirm,
                recover=args.recover,
                maximum_wait_seconds=args.maximum_wait,
            )
        elif args.subcommand == "explore":
            api_url, api_key, model = _vision_settings(args)
            result = interaction_learning.explore_application(
                knowledge.app_dir(args.app_id),
                api_url,
                api_key,
                model,
                __version__,
                policy=args.policy,
                max_actions=args.max_actions,
                confirm=args.confirm,
                max_depth=args.max_depth,
            )
        elif args.subcommand == "interactions":
            result = knowledge.list_interactions(
                knowledge.app_dir(args.app_id),
                args.command,
                args.limit,
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
