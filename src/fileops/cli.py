from __future__ import annotations

import argparse
import json
import logging
import sys
from pathlib import Path

from .models import RunReport
from .operations import CommonOptions, copy_items, delete_items, expand_sources, move_items, rename_items
from .reporting import print_summary, write_report


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="fileops",
        description="Safe file operation toolkit: copy, move, rename, delete.",
    )

    parser.add_argument("--log-level", default="INFO", choices=["DEBUG", "INFO", "WARNING", "ERROR"], help="Log verbosity.")
    parser.add_argument("--log-json", action="store_true", help="Output logs in JSON format.")

    subparsers = parser.add_subparsers(dest="command")

    copy_parser = subparsers.add_parser("copy", help="Copy files or directories.")
    _add_common_arguments(copy_parser, include_overwrite=True)
    copy_parser.add_argument("sources", nargs="+", help="Source paths or glob patterns.")
    copy_parser.add_argument("--dest", required=True, help="Destination path.")

    move_parser = subparsers.add_parser("move", help="Move files or directories.")
    _add_common_arguments(move_parser, include_overwrite=True)
    move_parser.add_argument("sources", nargs="+", help="Source paths or glob patterns.")
    move_parser.add_argument("--dest", required=True, help="Destination path.")

    rename_parser = subparsers.add_parser("rename", help="Rename files or directories by pattern.")
    _add_common_arguments(rename_parser, include_overwrite=True)
    rename_parser.add_argument("sources", nargs="+", help="Source paths or glob patterns.")
    rename_parser.add_argument("--pattern", required=True, help="Pattern with fields: {stem} {ext} {index} {parent} {date} {name}")
    rename_parser.add_argument("--start-index", type=int, default=1, help="Start index used in pattern.")

    delete_parser = subparsers.add_parser("delete", help="Delete files or directories.")
    _add_common_arguments(delete_parser, include_overwrite=False)
    delete_parser.add_argument("sources", nargs="+", help="Source paths or glob patterns.")
    delete_parser.add_argument("--trash", dest="use_trash", action="store_true", help="Move targets to system trash.")
    delete_parser.add_argument("--hard", dest="use_trash", action="store_false", help="Permanently delete targets.")
    delete_parser.set_defaults(use_trash=True)

    return parser


def _add_common_arguments(parser: argparse.ArgumentParser, include_overwrite: bool) -> None:
    parser.add_argument("--workspace", default=".", help="Workspace root. All paths must stay inside this root.")
    parser.add_argument("--dry-run", action="store_true", help="Preview actions without changing files.")
    parser.add_argument("--yes", action="store_true", help="Skip interactive confirmation prompts.")
    parser.add_argument("--recursive", action="store_true", help="Enable recursive glob expansion (**).")
    parser.add_argument("--report", help="Write execution report to a JSON file.")
    if include_overwrite:
        parser.add_argument(
            "--overwrite",
            default="never",
            choices=["never", "always", "rename"],
            help="Conflict policy when destination already exists.",
        )


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    if not args.command:
        parser.print_help()
        return 1

    _configure_logging(args.log_level, args.log_json)

    workspace = Path(args.workspace).resolve(strict=False)
    report = RunReport(command=args.command, dry_run_mode=bool(args.dry_run), workspace=str(workspace))

    sources = expand_sources(args.sources, recursive=bool(args.recursive), include_dirs=True)
    if not sources:
        logging.error("No source paths matched the given input.")
        return 1

    try:
        if args.command in {"copy", "move", "rename"}:
            _confirm_if_needed(args, sources_count=len(sources))
            common = CommonOptions(workspace=workspace, dry_run=bool(args.dry_run), overwrite=args.overwrite)

            if args.command == "copy":
                destination = Path(args.dest).resolve(strict=False)
                results = copy_items(sources, destination, common)
            elif args.command == "move":
                destination = Path(args.dest).resolve(strict=False)
                results = move_items(sources, destination, common)
            else:
                results = rename_items(sources, args.pattern, args.start_index, common)

        else:
            _confirm_if_needed(args, sources_count=len(sources))
            results = delete_items(
                sources=sources,
                workspace=workspace,
                dry_run=bool(args.dry_run),
                use_trash=bool(args.use_trash),
            )

    except ValueError as exc:
        logging.error(str(exc))
        return 2
    except RuntimeError as exc:
        logging.error(str(exc))
        return 2

    for item in results:
        report.add(item)
        level = logging.INFO if item.status.value != "failed" else logging.ERROR
        logging.log(level, "%s | %s -> %s | %s", item.operation, item.source, item.destination, item.message)

    print_summary(report)

    report_path = Path(args.report).resolve(strict=False) if args.report else None
    output_path = write_report(report, report_path)
    if output_path is not None:
        logging.info("Report written to %s", output_path)

    summary = report.summary()
    return 2 if summary["failed"] > 0 else 0


def _confirm_if_needed(args: argparse.Namespace, sources_count: int) -> None:
    if args.yes or args.dry_run:
        return

    if args.command == "delete":
        action = "trash" if args.use_trash else "hard-delete"
        prompt = f"About to {action} {sources_count} paths. Continue? [y/N]: "
    else:
        prompt = f"About to execute '{args.command}' for {sources_count} paths. Continue? [y/N]: "

    user_input = input(prompt).strip().lower()
    if user_input not in {"y", "yes"}:
        raise RuntimeError("Operation canceled by user.")


def _configure_logging(level: str, log_json: bool) -> None:
    root_logger = logging.getLogger()
    root_logger.setLevel(level)
    root_logger.handlers.clear()

    handler = logging.StreamHandler(stream=sys.stderr)
    if log_json:
        handler.setFormatter(_JsonFormatter())
    else:
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))

    root_logger.addHandler(handler)


class _JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        payload = {
            "time": self.formatTime(record, datefmt="%Y-%m-%dT%H:%M:%S"),
            "level": record.levelname,
            "message": record.getMessage(),
            "logger": record.name,
        }
        return json.dumps(payload, ensure_ascii=False)


if __name__ == "__main__":
    raise SystemExit(main())
