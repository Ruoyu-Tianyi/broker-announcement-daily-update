#!/usr/bin/env python3
"""Initialize and safely maintain a dated broker-announcement workspace."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import sys
from datetime import date, datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo


CONFIG_NAME = "workspace.json"
MASTER_NAME = "【日更】券商公告.xlsx"
RUN_PATTERN = re.compile(r"^run-(\d{3})$")


class WorkspaceError(RuntimeError):
    pass


def now_iso(timezone: str) -> str:
    return datetime.now(ZoneInfo(timezone)).isoformat(timespec="seconds")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def atomic_json(path: Path, payload: dict[str, Any]) -> None:
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temp, path)


def relative_to_root(root: Path, path: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError as error:
        raise WorkspaceError(f"Path must stay inside workspace: {path}") from error


def load_config(root: Path) -> dict[str, Any]:
    config_path = root / CONFIG_NAME
    if not config_path.is_file():
        raise WorkspaceError("Workspace is not initialized; workspace.json is missing")
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise WorkspaceError(f"Cannot read workspace.json: {error}") from error
    if config.get("schema_version") != 1 or not config.get("master_relative_path"):
        raise WorkspaceError("Unsupported or incomplete workspace.json")
    return config


def master_path(root: Path, config: dict[str, Any]) -> Path:
    path = (root / config["master_relative_path"]).resolve()
    relative_to_root(root, path)
    if not path.is_file():
        raise WorkspaceError(f"Active master workbook is missing: {path}")
    return path


def validate_batch_date(raw: str) -> str:
    try:
        return date.fromisoformat(raw).isoformat()
    except ValueError as error:
        raise WorkspaceError("Batch date must use YYYY-MM-DD") from error


def initialize(root: Path, source_master: Path, timezone: str) -> dict[str, Any]:
    root = root.resolve()
    source_master = source_master.resolve()
    if (root / CONFIG_NAME).exists():
        raise WorkspaceError(f"Workspace is already initialized: {root / CONFIG_NAME}")
    if not source_master.is_file() or source_master.suffix.lower() != ".xlsx":
        raise WorkspaceError("Initial master must be an existing .xlsx file")
    ZoneInfo(timezone)
    for name in ("master", "input", "output", "archive"):
        candidate = root / name
        if candidate.exists() and not candidate.is_dir():
            raise WorkspaceError(f"Workspace path conflicts with a file: {candidate}")
        candidate.mkdir(parents=True, exist_ok=True)
    (root / "archive" / "master").mkdir(parents=True, exist_ok=True)
    active_master = root / "master" / MASTER_NAME
    if active_master.exists():
        raise WorkspaceError(f"Refusing to overwrite existing active master: {active_master}")
    shutil.copy2(source_master, active_master)
    master_hash = sha256(active_master)
    config = {
        "schema_version": 1,
        "created_at": now_iso(timezone),
        "timezone": timezone,
        "master_relative_path": relative_to_root(root, active_master),
        "active_master": {
            "sha256": master_hash,
            "batch_date": None,
            "updated_at": None,
        },
        "initial_master_source": str(source_master),
    }
    atomic_json(root / CONFIG_NAME, config)
    return {
        "status": "initialized",
        "root": str(root),
        "active_master": str(active_master),
        "master_sha256": master_hash,
    }


def next_run_id(date_input_root: Path) -> str:
    numbers = []
    if date_input_root.is_dir():
        for child in date_input_root.iterdir():
            match = RUN_PATTERN.fullmatch(child.name)
            if match and child.is_dir():
                numbers.append(int(match.group(1)))
    return f"run-{(max(numbers, default=0) + 1):03d}"


def start_run(root: Path, batch_date: str, sources: list[Path]) -> dict[str, Any]:
    root = root.resolve()
    config = load_config(root)
    batch_date = validate_batch_date(batch_date)
    if len(sources) != 2:
        raise WorkspaceError("A daily run requires exactly two source Excel files")
    resolved_sources = [item.resolve() for item in sources]
    for source in resolved_sources:
        if not source.is_file() or source.suffix.lower() != ".xlsx":
            raise WorkspaceError(f"Source is not an existing .xlsx file: {source}")
    source_hashes = [sha256(item) for item in resolved_sources]
    source_signature = sorted(source_hashes)
    date_input_root = root / "input" / batch_date
    if date_input_root.is_dir():
        for run_manifest in date_input_root.glob("run-*/run.json"):
            try:
                existing = json.loads(run_manifest.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            existing_hashes = [item.get("sha256") for item in existing.get("source_excels", [])]
            if sorted(existing_hashes) == source_signature:
                raise WorkspaceError(
                    f"The same source pair is already archived in {run_manifest.parent}"
                )
    run_id = next_run_id(date_input_root)
    input_run = date_input_root / run_id
    output_run = root / "output" / batch_date / run_id
    input_run.mkdir(parents=True, exist_ok=False)
    output_run.mkdir(parents=True, exist_ok=False)
    archived_sources = []
    for index, (source, source_hash) in enumerate(zip(resolved_sources, source_hashes), start=1):
        archived = input_run / f"{index:02d}_{source.name}"
        shutil.copy2(source, archived)
        archived_sources.append({
            "original_path": str(source),
            "archived_path": relative_to_root(root, archived),
            "sha256": source_hash,
            "bytes": archived.stat().st_size,
        })
    active_master = master_path(root, config)
    manifest = {
        "schema_version": 1,
        "batch_date": batch_date,
        "run_id": run_id,
        "status": "initialized",
        "created_at": now_iso(config["timezone"]),
        "master_before": {
            "path": relative_to_root(root, active_master),
            "sha256": sha256(active_master),
        },
        "source_excels": archived_sources,
        "pdfs": [],
        "output_directory": relative_to_root(root, output_run),
    }
    atomic_json(input_run / "run.json", manifest)
    return {
        "status": "initialized",
        "batch_date": batch_date,
        "run_id": run_id,
        "input_directory": str(input_run),
        "output_directory": str(output_run),
        "active_master": str(active_master),
        "manifest": str(input_run / "run.json"),
    }


def attach_pdfs(
    root: Path, batch_date: str, run_id: str, pdfs: list[Path]
) -> dict[str, Any]:
    root = root.resolve()
    config = load_config(root)
    batch_date = validate_batch_date(batch_date)
    manifest_path = root / "input" / batch_date / run_id / "run.json"
    if not manifest_path.is_file():
        raise WorkspaceError(f"Run manifest does not exist: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    pdf_dir = manifest_path.parent / "pdf"
    pdf_dir.mkdir(exist_ok=True)
    existing_hashes = {item["sha256"] for item in manifest.get("pdfs", [])}
    added = []
    for source in (item.resolve() for item in pdfs):
        if not source.is_file() or source.suffix.lower() != ".pdf":
            raise WorkspaceError(f"Source is not an existing PDF: {source}")
        source_hash = sha256(source)
        if source_hash in existing_hashes:
            continue
        target = pdf_dir / f"{len(manifest.get('pdfs', [])) + len(added) + 1:02d}_{source.name}"
        if target.exists():
            raise WorkspaceError(f"Refusing to overwrite archived PDF: {target}")
        shutil.copy2(source, target)
        record = {
            "original_path": str(source),
            "archived_path": relative_to_root(root, target),
            "sha256": source_hash,
            "bytes": target.stat().st_size,
        }
        added.append(record)
        existing_hashes.add(source_hash)
    manifest.setdefault("pdfs", []).extend(added)
    manifest["pdfs_updated_at"] = now_iso(config["timezone"])
    atomic_json(manifest_path, manifest)
    return {"status": "updated", "added_count": len(added), "pdf_directory": str(pdf_dir)}


def promote(
    root: Path,
    batch_date: str,
    run_id: str,
    final_workbook: Path,
    verification_report: Path,
) -> dict[str, Any]:
    root = root.resolve()
    config = load_config(root)
    batch_date = validate_batch_date(batch_date)
    manifest_path = root / "input" / batch_date / run_id / "run.json"
    if not manifest_path.is_file():
        raise WorkspaceError(f"Run manifest does not exist: {manifest_path}")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    final_workbook = final_workbook.resolve()
    verification_report = verification_report.resolve()
    output_root = (root / manifest["output_directory"]).resolve()
    relative_to_root(output_root, final_workbook)
    relative_to_root(output_root, verification_report)
    if not final_workbook.is_file() or final_workbook.suffix.lower() != ".xlsx":
        raise WorkspaceError("Final workbook must be an existing .xlsx inside the run output directory")
    try:
        report = json.loads(verification_report.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise WorkspaceError(f"Cannot read verification report: {error}") from error
    if report.get("status") != "complete":
        raise WorkspaceError("Verification report is not complete")
    if report.get("unresolved_count", 0) != 0:
        raise WorkspaceError("Cannot promote a workbook with unresolved decisions")
    candidates = report.get("board_resolution_candidates")
    if isinstance(candidates, list) and candidates:
        raise WorkspaceError("Stage-one workbook still requires PDFs and cannot become the master")

    active_master = master_path(root, config)
    current_hash = sha256(active_master)
    if current_hash != manifest["master_before"]["sha256"]:
        raise WorkspaceError("Active master changed after this run started; refusing stale promotion")
    archive_dir = root / "archive" / "master" / batch_date / run_id
    archive_dir.mkdir(parents=True, exist_ok=False)
    archived_master = archive_dir / f"{active_master.stem}_before_{batch_date}_{run_id}.xlsx"
    shutil.copy2(active_master, archived_master)
    final_hash = sha256(final_workbook)
    staged_master = active_master.with_suffix(active_master.suffix + ".promote.tmp")
    shutil.copy2(final_workbook, staged_master)
    if sha256(staged_master) != final_hash:
        raise WorkspaceError("Staged master hash does not match final workbook")
    os.replace(staged_master, active_master)
    if sha256(active_master) != final_hash:
        raise WorkspaceError("Promoted master hash verification failed")

    manifest.update({
        "status": "complete",
        "completed_at": now_iso(config["timezone"]),
        "final_workbook": relative_to_root(root, final_workbook),
        "verification_report": relative_to_root(root, verification_report),
        "final_sha256": final_hash,
        "archived_master": relative_to_root(root, archived_master),
    })
    atomic_json(manifest_path, manifest)
    config["active_master"] = {
        "sha256": final_hash,
        "batch_date": batch_date,
        "run_id": run_id,
        "updated_at": now_iso(config["timezone"]),
    }
    atomic_json(root / CONFIG_NAME, config)
    return {
        "status": "complete",
        "active_master": str(active_master),
        "active_master_sha256": final_hash,
        "archived_previous_master": str(archived_master),
        "run_manifest": str(manifest_path),
    }


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    init = subparsers.add_parser("init")
    init.add_argument("--root", type=Path, required=True)
    init.add_argument("--master", type=Path, required=True)
    init.add_argument("--timezone", default="Asia/Shanghai")
    start = subparsers.add_parser("start-run")
    start.add_argument("--root", type=Path, required=True)
    start.add_argument("--date", required=True)
    start.add_argument("--source", type=Path, action="append", required=True)
    attach = subparsers.add_parser("attach-pdf")
    attach.add_argument("--root", type=Path, required=True)
    attach.add_argument("--date", required=True)
    attach.add_argument("--run", required=True)
    attach.add_argument("--pdf", type=Path, action="append", required=True)
    finish = subparsers.add_parser("promote")
    finish.add_argument("--root", type=Path, required=True)
    finish.add_argument("--date", required=True)
    finish.add_argument("--run", required=True)
    finish.add_argument("--final", type=Path, required=True)
    finish.add_argument("--report", type=Path, required=True)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        if args.command == "init":
            result = initialize(args.root, args.master, args.timezone)
        elif args.command == "start-run":
            result = start_run(args.root, args.date, args.source)
        elif args.command == "attach-pdf":
            result = attach_pdfs(args.root, args.date, args.run, args.pdf)
        else:
            result = promote(args.root, args.date, args.run, args.final, args.report)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0
    except (WorkspaceError, OSError, ValueError) as error:
        print(json.dumps({"status": "error", "error": str(error)}, ensure_ascii=False, indent=2))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
