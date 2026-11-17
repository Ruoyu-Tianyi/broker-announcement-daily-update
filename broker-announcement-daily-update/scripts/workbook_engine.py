#!/usr/bin/env python3
"""File-level engine for the broker announcement daily update skill."""

from __future__ import annotations

import argparse
import copy
import json
import posixpath
import re
import shutil
import tempfile
import unicodedata
import warnings
import zipfile
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.datetime import from_excel
from openpyxl.worksheet.cell_range import CellRange


MIN_OPENPYXL = (3, 1, 5)
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
SHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DATE_HEADERS = {"见网日期", "挂网日期"}
BROKER_HEADERS = {"券商名称"}
TITLE_HEADERS = {"公告题目", "公告标题"}
RISKY_PACKAGE_PREFIXES = (
    "xl/activeX/",
    "xl/connections",
    "xl/ctrlProps/",
    "xl/customXml/",
    "xl/embeddings/",
    "xl/externalLinks/",
    "xl/pivotCache/",
    "xl/pivotTables/",
    "xl/richData/",
    "xl/slicers/",
    "xl/slicerCaches/",
    "xl/threadedComments/",
)


class EngineError(RuntimeError):
    pass


@dataclass(frozen=True)
class Record:
    source: str
    source_row: int
    seen_date: date | datetime
    broker: str
    title: str


@dataclass
class Block:
    title_row: int
    header_row: int
    end_row: int
    date_col: int
    broker_col: int
    title_col: int
    max_date: date | datetime | None
    date_sequence: list[date | datetime]
    record_rows: list[int]


def norm(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value))


def date_key(value: date | datetime | None) -> tuple[int, int, int]:
    if value is None:
        return (0, 0, 0)
    return (value.year, value.month, value.day)


def as_date(value: Any, epoch: datetime) -> date | datetime | None:
    if isinstance(value, (date, datetime)):
        return value
    if isinstance(value, (int, float)):
        converted = from_excel(value, epoch)
        return converted
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(text, fmt)
            except ValueError:
                pass
    return None


def json_value(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, CellRichText):
        return str(value)
    return value


def color_sig(color: Any) -> Any:
    if color is None:
        return None
    def stable(value: Any) -> Any:
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        return str(value)
    return {
        "type": stable(color.type),
        "rgb": stable(color.rgb),
        "indexed": stable(color.indexed),
        "theme": stable(color.theme),
        "tint": stable(color.tint),
        "auto": stable(color.auto),
    }


def side_sig(side: Any) -> Any:
    if side is None:
        return None
    return {"style": side.style, "color": color_sig(side.color)}


def style_sig(cell: Any) -> dict[str, Any]:
    font = cell.font
    fill = cell.fill
    border = cell.border
    alignment = cell.alignment
    protection = cell.protection
    return {
        "font": {
            "name": font.name,
            "sz": font.sz,
            "b": font.b,
            "i": font.i,
            "u": font.u,
            "strike": font.strike,
            "color": color_sig(font.color),
            "vertAlign": font.vertAlign,
            "charset": font.charset,
            "family": font.family,
            "scheme": font.scheme,
            "outline": font.outline,
            "shadow": font.shadow,
            "condense": font.condense,
            "extend": font.extend,
        },
        "fill": {
            "fill_type": fill.fill_type,
            "fg": color_sig(fill.fgColor),
            "bg": color_sig(fill.bgColor),
        },
        "border": {
            "left": side_sig(border.left),
            "right": side_sig(border.right),
            "top": side_sig(border.top),
            "bottom": side_sig(border.bottom),
            "diagonal": side_sig(border.diagonal),
            "vertical": side_sig(border.vertical),
            "horizontal": side_sig(border.horizontal),
            "diagonalUp": border.diagonalUp,
            "diagonalDown": border.diagonalDown,
            "outline": border.outline,
        },
        "alignment": {
            "horizontal": alignment.horizontal,
            "vertical": alignment.vertical,
            "textRotation": alignment.textRotation,
            "wrapText": alignment.wrapText,
            "shrinkToFit": alignment.shrinkToFit,
            "indent": alignment.indent,
            "relativeIndent": alignment.relativeIndent,
            "justifyLastLine": alignment.justifyLastLine,
            "readingOrder": alignment.readingOrder,
        },
        "number_format": cell.number_format,
        "protection": {"locked": protection.locked, "hidden": protection.hidden},
        "quotePrefix": getattr(cell, "quotePrefix", None),
        "pivotButton": getattr(cell, "pivotButton", None),
    }


def rich_value_sig(value: Any) -> Any:
    if not isinstance(value, CellRichText):
        return json_value(value)
    runs = []
    for item in value:
        if isinstance(item, TextBlock):
            runs.append({
                "text": item.text,
                "bold": item.font.b,
                "italic": item.font.i,
                "name": item.font.rFont,
                "size": item.font.sz,
            })
        else:
            runs.append({"text": str(item), "bold": None})
    return {"rich_text": runs}


def cell_sig(cell: Any | None) -> Any:
    if cell is None:
        return None
    hyperlink = cell.hyperlink
    comment = cell.comment
    signature = {
        "value": rich_value_sig(cell.value),
        "data_type": cell.data_type,
        "style": style_sig(cell),
        "hyperlink": None if hyperlink is None else {
            "target": hyperlink.target,
            "location": hyperlink.location,
            "tooltip": hyperlink.tooltip,
            "display": hyperlink.display,
        },
        "comment": None if comment is None else {
            "text": comment.text,
            "author": comment.author,
            "height": comment.height,
            "width": comment.width,
        },
    }
    if (
        signature["value"] is None
        and not cell.has_style
        and signature["hyperlink"] is None
        and signature["comment"] is None
    ):
        return None
    return signature


def row_dim_sig(dim: Any) -> dict[str, Any]:
    return {
        "height": dim.height,
        "hidden": dim.hidden,
        "outlineLevel": dim.outlineLevel,
        "collapsed": dim.collapsed,
        "thickTop": dim.thickTop,
        "thickBottom": dim.thickBot,
        "style": dim.style,
    }


def col_dim_sig(dim: Any) -> dict[str, Any]:
    return {
        "width": dim.width,
        "hidden": dim.hidden,
        "bestFit": dim.bestFit,
        "min": dim.min,
        "max": dim.max,
        "outlineLevel": dim.outlineLevel,
        "collapsed": dim.collapsed,
        "style": dim.style,
    }


def snapshot_sheet(ws: Any, start_row: int = 1) -> dict[str, Any]:
    max_row = ws.max_row
    max_col = ws.max_column
    cells: dict[str, Any] = {}
    for row in range(start_row, max_row + 1):
        for col in range(1, max_col + 1):
            sig = cell_sig(ws._cells.get((row, col)))
            if sig is not None:
                cells[f"{row - start_row}:{col}"] = sig
    row_dims = {
        str(idx - start_row): row_dim_sig(dim)
        for idx, dim in ws.row_dimensions.items()
        if idx >= start_row
    }
    merges = []
    for merged in ws.merged_cells.ranges:
        if merged.max_row < start_row:
            continue
        if merged.min_row < start_row:
            merges.append(f"CROSS:{merged}")
        else:
            merges.append(
                f"{merged.min_row-start_row}:{merged.min_col}:"
                f"{merged.max_row-start_row}:{merged.max_col}"
            )
    return {
        "span": [max_row - start_row + 1, max_col],
        "cells": cells,
        "row_dimensions": row_dims,
        "merges": sorted(merges),
        "columns": {key: col_dim_sig(dim) for key, dim in ws.column_dimensions.items()},
        "freeze_panes": str(ws.freeze_panes) if ws.freeze_panes else None,
        "sheet_view": {
            "showGridLines": ws.sheet_view.showGridLines,
            "zoomScale": ws.sheet_view.zoomScale,
            "rightToLeft": ws.sheet_view.rightToLeft,
        },
        "auto_filter": ws.auto_filter.ref,
        "sheet_format": {
            "defaultRowHeight": ws.sheet_format.defaultRowHeight,
            "defaultColWidth": ws.sheet_format.defaultColWidth,
        },
        "print": {
            "orientation": ws.page_setup.orientation,
            "paperSize": ws.page_setup.paperSize,
            "fitToWidth": ws.page_setup.fitToWidth,
            "fitToHeight": ws.page_setup.fitToHeight,
            "margins": {
                "left": ws.page_margins.left,
                "right": ws.page_margins.right,
                "top": ws.page_margins.top,
                "bottom": ws.page_margins.bottom,
                "header": ws.page_margins.header,
                "footer": ws.page_margins.footer,
            },
        },
    }


def workbook_sheet_snapshots(wb: Any, exclude: str | None = None) -> dict[str, Any]:
    return {
        ws.title: snapshot_sheet(ws)
        for ws in wb.worksheets
        if ws.title != exclude
    }


def first_difference(left: Any, right: Any, path: str = "root") -> str | None:
    if type(left) is not type(right):
        return f"{path}: type {type(left).__name__} != {type(right).__name__}"
    if isinstance(left, dict):
        if set(left) != set(right):
            missing = sorted(set(left) - set(right))[:5]
            added = sorted(set(right) - set(left))[:5]
            return f"{path}: keys missing={missing}, added={added}"
        for key in sorted(left):
            difference = first_difference(left[key], right[key], f"{path}.{key}")
            if difference:
                return difference
        return None
    if isinstance(left, list):
        if len(left) != len(right):
            return f"{path}: length {len(left)} != {len(right)}"
        for index, (left_item, right_item) in enumerate(zip(left, right)):
            difference = first_difference(left_item, right_item, f"{path}[{index}]")
            if difference:
                return difference
        return None
    if left != right:
        return f"{path}: {left!r} != {right!r}"
    return None


def load_xlsx(path: Path) -> tuple[Any, list[str]]:
    caught: list[str] = []
    with warnings.catch_warnings(record=True) as records:
        warnings.simplefilter("always")
        wb = load_workbook(path, data_only=False, keep_links=True, rich_text=True)
        caught.extend(str(item.message) for item in records)
    return wb, caught


def package_entries(path: Path) -> set[str]:
    with zipfile.ZipFile(path) as archive:
        return set(archive.namelist())


def risky_parts(path: Path) -> list[str]:
    entries = package_entries(path)
    risky = [entry for entry in entries if entry.startswith(RISKY_PACKAGE_PREFIXES)]
    risky.extend(entry for entry in entries if entry.endswith("vbaProject.bin"))
    return sorted(set(risky))


def header_mapping(ws: Any, row: int) -> dict[str, int] | None:
    found: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        label = norm(ws.cell(row, col).value)
        if label in DATE_HEADERS:
            found["date"] = col
        elif label in BROKER_HEADERS:
            found["broker"] = col
        elif label in TITLE_HEADERS:
            found["title"] = col
    return found if set(found) == {"date", "broker", "title"} else None


def find_blocks(ws: Any, epoch: datetime) -> list[Block]:
    headers: list[tuple[int, dict[str, int]]] = []
    for row in range(1, ws.max_row + 1):
        mapping = header_mapping(ws, row)
        if mapping:
            headers.append((row, mapping))
    blocks: list[Block] = []
    for idx, (header_row, mapping) in enumerate(headers):
        title_row = header_row - 1
        if title_row < 1 or "券商公告" not in norm(ws.cell(title_row, mapping["date"]).value):
            continue
        end_row = (headers[idx + 1][0] - 2) if idx + 1 < len(headers) else ws.max_row
        current_date = None
        dates: list[date | datetime] = []
        distinct_dates: list[date | datetime] = []
        record_rows: list[int] = []
        for row in range(header_row + 1, end_row + 1):
            raw_date = ws.cell(row, mapping["date"]).value
            parsed = as_date(raw_date, epoch)
            if parsed is not None:
                current_date = parsed
                if not distinct_dates or date_key(distinct_dates[-1]) != date_key(parsed):
                    distinct_dates.append(parsed)
            title = ws.cell(row, mapping["title"]).value
            if title not in (None, ""):
                record_rows.append(row)
                if current_date is not None:
                    dates.append(current_date)
        blocks.append(Block(
            title_row=title_row,
            header_row=header_row,
            end_row=end_row,
            date_col=mapping["date"],
            broker_col=mapping["broker"],
            title_col=mapping["title"],
            max_date=max(dates, key=date_key) if dates else None,
            date_sequence=distinct_dates,
            record_rows=record_rows,
        ))
    return blocks


def choose_target_sheet_and_block(wb: Any) -> tuple[Any, Block, list[dict[str, Any]]]:
    candidates: list[tuple[Any, Block]] = []
    for ws in wb.worksheets:
        for block in find_blocks(ws, wb.epoch):
            if block.record_rows and block.max_date is not None:
                candidates.append((ws, block))
    if not candidates:
        raise EngineError("No complete daily announcement block with dated records was found")
    candidates.sort(key=lambda pair: (date_key(pair[1].max_date), -pair[1].title_row), reverse=True)
    ws, selected = candidates[0]
    summary = [{
        "sheet": item_ws.title,
        "title_row": item.title_row,
        "header_row": item.header_row,
        "records": len(item.record_rows),
        "max_date": json_value(item.max_date),
    } for item_ws, item in candidates]
    return ws, selected, summary


def extract_source(path: Path) -> tuple[list[Record], dict[str, Any]]:
    wb, load_warnings = load_xlsx(path)
    matches: list[tuple[Any, int, dict[str, int]]] = []
    for ws in wb.worksheets:
        for row in range(1, ws.max_row + 1):
            mapping = header_mapping(ws, row)
            if mapping:
                matches.append((ws, row, mapping))
    if len(matches) != 1:
        raise EngineError(f"{path.name}: expected exactly one source header, found {len(matches)}")
    ws, header_row, mapping = matches[0]
    records: list[Record] = []
    current_date = None
    current_broker = None
    for row in range(header_row + 1, ws.max_row + 1):
        title_raw = ws.cell(row, mapping["title"]).value
        if title_raw in (None, ""):
            continue
        title = str(title_raw).strip()
        if norm(title) in {"合计", "总计"}:
            continue
        parsed_date = as_date(ws.cell(row, mapping["date"]).value, wb.epoch)
        if parsed_date is not None:
            current_date = parsed_date
        broker_raw = ws.cell(row, mapping["broker"]).value
        if broker_raw not in (None, ""):
            current_broker = str(broker_raw).strip()
        if current_date is None or not current_broker:
            raise EngineError(f"{path.name}!{ws.title} row {row}: missing effective date or broker")
        records.append(Record(path.name, row, current_date, current_broker, title))
    duplicates: list[dict[str, Any]] = []
    seen: dict[tuple[Any, str, str], int] = {}
    for record in records:
        key = (date_key(record.seen_date), record.broker, record.title)
        if key in seen:
            duplicates.append({"first_row": seen[key], "duplicate_row": record.source_row, "title": record.title})
        else:
            seen[key] = record.source_row
    return records, {
        "file": str(path),
        "sheet": ws.title,
        "header_row": header_row,
        "record_count": len(records),
        "max_date": json_value(max((r.seen_date for r in records), key=date_key)) if records else None,
        "duplicates": duplicates,
        "warnings": load_warnings,
    }


def workbook_blockers(path: Path, wb: Any, target_ws: Any, insert_at: int) -> list[str]:
    blockers: list[str] = []
    if path.suffix.lower() != ".xlsx":
        blockers.append("Only .xlsx is supported by the packaged engine")
    for part in risky_parts(path):
        blockers.append(f"Unsupported OOXML part: {part}")
    if getattr(wb, "_external_links", None):
        blockers.append("Workbook contains external links")
    defined_names = list(wb.defined_names.values())
    if defined_names:
        blockers.append(f"Workbook contains {len(defined_names)} defined name(s)")
    for ws in wb.worksheets:
        if ws.tables:
            blockers.append(f"{ws.title}: contains structured table(s)")
        if len(ws.conditional_formatting):
            blockers.append(f"{ws.title}: contains conditional formatting")
        if ws.data_validations.count:
            blockers.append(f"{ws.title}: contains data validation")
        if ws._charts or ws._images:
            blockers.append(f"{ws.title}: contains chart/image drawings")
        if ws.auto_filter.ref:
            blockers.append(f"{ws.title}: contains an active autofilter")
        if len(ws.row_breaks.brk) or len(ws.col_breaks.brk):
            blockers.append(f"{ws.title}: contains manual page breaks")
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    blockers.append(f"{ws.title}: contains formula at {cell.coordinate}")
                    break
            if blockers and blockers[-1].startswith(f"{ws.title}: contains formula"):
                break
    for merged in target_ws.merged_cells.ranges:
        if merged.min_row < insert_at <= merged.max_row:
            blockers.append(f"Merged range crosses insertion point: {merged}")
    return sorted(set(blockers))


def preflight(master: Path, sources: list[Path]) -> dict[str, Any]:
    version = tuple(int(part) for part in openpyxl.__version__.split(".")[:3])
    if version < MIN_OPENPYXL:
        raise EngineError(f"openpyxl {openpyxl.__version__} is too old; need >= 3.1.5")
    master_wb, master_warnings = load_xlsx(master)
    target_ws, block, candidates = choose_target_sheet_and_block(master_wb)
    row_above = block.title_row - 1
    separator = row_above >= 1 and all(
        target_ws.cell(row_above, col).value in (None, "")
        for col in range(1, target_ws.max_column + 1)
    )
    insert_at = row_above if separator else block.title_row
    source_reports = []
    all_records: list[list[Record]] = []
    for source in sources:
        records, report = extract_source(source)
        all_records.append(records)
        source_reports.append(report)
    blockers = workbook_blockers(master, master_wb, target_ws, insert_at)
    return {
        "engine": {"openpyxl": openpyxl.__version__, "rich_text": True},
        "master": str(master),
        "master_warnings": master_warnings,
        "target_sheet": target_ws.title,
        "selected_block": {
            "title_row": block.title_row,
            "header_row": block.header_row,
            "end_row": block.end_row,
            "max_date": json_value(block.max_date),
            "record_count": len(block.record_rows),
            "date_order": infer_date_order(block.date_sequence),
        },
        "candidate_blocks": candidates,
        "separator_row": row_above if separator else None,
        "insert_at": insert_at,
        "sources": source_reports,
        "source_record_count": sum(len(records) for records in all_records),
        "insert_count": 2 + sum(len(records) for records in all_records),
        "package_entries": sorted(package_entries(master)),
        "blockers": blockers,
        "status": "ready" if not blockers else "blocked",
    }


def infer_date_order(sequence: list[date | datetime]) -> str:
    keys = [date_key(item) for item in sequence]
    if len(keys) < 2:
        return "undetermined"
    if keys == sorted(keys):
        return "ascending"
    if keys == sorted(keys, reverse=True):
        return "descending"
    return "mixed"


def shift_row_dimensions(ws: Any, start: int, amount: int) -> None:
    items = [(idx, copy.copy(dim)) for idx, dim in ws.row_dimensions.items() if idx >= start]
    for idx, _ in items:
        del ws.row_dimensions[idx]
    for idx, dim in sorted(items, reverse=True):
        dim.index = idx + amount
        ws.row_dimensions[idx + amount] = dim


def insert_rows_preserving_merges(ws: Any, insert_at: int, amount: int) -> list[CellRange]:
    original_merges = [CellRange(str(item)) for item in ws.merged_cells.ranges]
    annotations = [
        (row, col, copy.copy(cell.hyperlink), copy.copy(cell.comment))
        for (row, col), cell in ws._cells.items()
        if row >= insert_at and (cell.hyperlink is not None or cell.comment is not None)
    ]
    crossing = [item for item in original_merges if item.min_row < insert_at <= item.max_row]
    if crossing:
        raise EngineError(f"Cannot insert through merged range(s): {', '.join(map(str, crossing))}")
    for item in original_merges:
        if item.min_row >= insert_at:
            ws.unmerge_cells(str(item))
    ws.insert_rows(insert_at, amount)
    shift_row_dimensions(ws, insert_at, amount)
    for row, col, hyperlink, comment in annotations:
        target = ws.cell(row + amount, col)
        target.hyperlink = hyperlink
        target.comment = comment
    for item in original_merges:
        if item.min_row >= insert_at:
            shifted = copy.copy(item)
            shifted.shift(row_shift=amount)
            ws.merge_cells(str(shifted))
    return original_merges


def clone_row_dimension(ws: Any, source_row: int, target_row: int) -> None:
    if source_row not in ws.row_dimensions:
        return
    dim = copy.copy(ws.row_dimensions[source_row])
    dim.index = target_row
    ws.row_dimensions[target_row] = dim


def display_width(value: Any) -> int:
    """Approximate the rendered width of a single-line title in CJK text units."""
    return sum(
        2 if unicodedata.east_asian_width(char) in {"W", "F", "A"} else 1
        for char in str(value or "")
    )


def match_row_height_to_precedent(
    ws: Any, title_col: int, target_row: int, precedent_rows: list[int]
) -> dict[str, Any]:
    """Match a new title row to the closest single-line row in the latest block."""
    target_text = str(ws.cell(target_row, title_col).value or "")
    target_width = display_width(target_text)
    candidates: list[tuple[int, int, int]] = []
    for row in precedent_rows:
        text = str(ws.cell(row, title_col).value or "")
        if not text or "\n" in text or "\r" in text:
            continue
        width = display_width(text)
        # Prefer a precedent at least as wide on an equal-distance tie so long
        # titles are not assigned a row height known only to fit shorter text.
        candidates.append((abs(width - target_width), 0 if width >= target_width else 1, row))
    if not candidates:
        raise EngineError("Latest daily block has no single-line row-height precedent")
    _, _, source_row = min(candidates)
    source_height = ws.row_dimensions[source_row].height
    ws.row_dimensions[target_row].height = source_height
    return {
        "row": target_row,
        "title_display_width": target_width,
        "precedent_row": source_row,
        "precedent_display_width": display_width(ws.cell(source_row, title_col).value),
        "height": source_height,
    }


def copy_cell(source: Any, target: Cell, include_value: bool = True) -> None:
    if include_value:
        target.value = copy.copy(source.value)
    if source.has_style:
        target._style = copy.copy(source._style)
    if source.number_format:
        target.number_format = source.number_format
    target.hyperlink = copy.copy(source.hyperlink)
    target.comment = copy.copy(source.comment)


def contiguous_groups(values: list[Any]) -> Iterable[tuple[int, int]]:
    if not values:
        return
    start = 0
    for idx in range(1, len(values) + 1):
        if idx == len(values) or values[idx] != values[start]:
            yield start, idx - 1
            start = idx


def is_board_resolution(title: str) -> bool:
    compact = norm(title)
    if any(term in compact for term in ("监事会", "股东会", "股东大会")):
        return False
    return bool(re.search(r"董事会.*决议(?:的)?公告", compact))


def restore_printer_settings(source: Path, output: Path) -> None:
    with zipfile.ZipFile(source) as src_zip, zipfile.ZipFile(output) as out_zip:
        data = {name: out_zip.read(name) for name in out_zip.namelist()}
        original_names = set(src_zip.namelist())
        rel_files = sorted(
            name for name in original_names
            if name.startswith("xl/worksheets/_rels/") and name.endswith(".xml.rels")
        )
        changed = False
        for rel_name in rel_files:
            src_rel_root = ET.fromstring(src_zip.read(rel_name))
            printer_rels = [
                rel for rel in src_rel_root
                if rel.attrib.get("Type", "").endswith("/printerSettings")
            ]
            if not printer_rels:
                continue
            sheet_part = rel_name.replace("/_rels/", "/")[:-5]
            if sheet_part not in data:
                raise EngineError(f"Output is missing worksheet part {sheet_part}")
            if rel_name in data:
                out_rel_root = ET.fromstring(data[rel_name])
            else:
                out_rel_root = ET.Element(f"{{{REL_NS}}}Relationships")
            used_ids = {rel.attrib.get("Id") for rel in out_rel_root}
            sheet_root = ET.fromstring(data[sheet_part])
            page_setup = sheet_root.find(f"{{{SHEET_NS}}}pageSetup")
            if page_setup is None:
                page_setup = ET.SubElement(sheet_root, f"{{{SHEET_NS}}}pageSetup")
            for rel in printer_rels:
                new_id = rel.attrib["Id"]
                if new_id in used_ids:
                    counter = 1
                    while f"rIdPrinter{counter}" in used_ids:
                        counter += 1
                    new_id = f"rIdPrinter{counter}"
                target = rel.attrib["Target"]
                source_dir = posixpath.dirname(sheet_part)
                target_part = posixpath.normpath(posixpath.join(source_dir, target))
                if target_part not in original_names:
                    raise EngineError(f"Missing printer settings payload {target_part}")
                data[target_part] = src_zip.read(target_part)
                ET.SubElement(out_rel_root, f"{{{REL_NS}}}Relationship", {
                    "Id": new_id,
                    "Type": rel.attrib["Type"],
                    "Target": target,
                })
                used_ids.add(new_id)
                page_setup.set(f"{{{R_NS}}}id", new_id)
                changed = True
            data[rel_name] = ET.tostring(out_rel_root, encoding="utf-8", xml_declaration=True)
            data[sheet_part] = ET.tostring(sheet_root, encoding="utf-8", xml_declaration=True)
        if changed:
            content_root = ET.fromstring(data["[Content_Types].xml"])
            if not any(
                item.attrib.get("Extension") == "bin"
                and "printerSettings" in item.attrib.get("ContentType", "")
                for item in content_root
            ):
                ET.SubElement(content_root, f"{{{CONTENT_NS}}}Default", {
                    "Extension": "bin",
                    "ContentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings",
                })
            # Some strict OOXML readers reject a prefixed `ns0:Types` root even
            # though it is namespace-equivalent.  Content_Types.xml conventionally
            # uses the package content-types namespace as the default namespace.
            ET.register_namespace("", CONTENT_NS)
            data["[Content_Types].xml"] = ET.tostring(
                content_root, encoding="utf-8", xml_declaration=True
            )
    if not changed:
        return
    temp = output.with_suffix(output.suffix + ".printer.tmp")
    with zipfile.ZipFile(temp, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, payload in data.items():
            archive.writestr(name, payload)
    temp.replace(output)


def verify_package(source: Path, output: Path) -> dict[str, Any]:
    before = package_entries(source)
    after = package_entries(output)
    lost = sorted(before - after)
    allowed_lost = {"xl/sharedStrings.xml", "xl/calcChain.xml"}
    unexpected_lost = [entry for entry in lost if entry not in allowed_lost]
    if unexpected_lost:
        raise EngineError(f"OOXML parts were lost: {unexpected_lost}")
    return {"lost_allowed": sorted(set(lost) & allowed_lost), "added": sorted(after - before)}


def sort_source_bundles(
    bundles: list[tuple[Path, list[Record], dict[str, Any]]], order: str
) -> list[tuple[Path, list[Record], dict[str, Any]]]:
    reverse = order == "descending"
    if order not in {"ascending", "descending"}:
        return bundles
    return sorted(
        bundles,
        key=lambda item: date_key(max((r.seen_date for r in item[1]), key=date_key)) if item[1] else (0, 0, 0),
        reverse=reverse,
    )


def stage1(master: Path, sources: list[Path], output: Path) -> dict[str, Any]:
    if output.exists():
        raise EngineError(f"Refusing to overwrite existing output: {output}")
    report = preflight(master, sources)
    if report["blockers"]:
        raise EngineError("Preflight blocked: " + "; ".join(report["blockers"]))
    wb, load_warnings = load_xlsx(master)
    ws, block, _ = choose_target_sheet_and_block(wb)
    insert_at = report["insert_at"]
    bundles = []
    for source in sources:
        records, source_report = extract_source(source)
        bundles.append((source, records, source_report))
    bundles = sort_source_bundles(bundles, report["selected_block"]["date_order"])
    records = [record for _, bundle_records, _ in bundles for record in bundle_records]
    k = 2 + len(records)
    before_history = snapshot_sheet(ws, insert_at)
    before_other_sheets = workbook_sheet_snapshots(wb, exclude=ws.title)
    before_sheet_state = ws.sheet_state
    original_merges = insert_rows_preserving_merges(ws, insert_at, k)

    shifted_title_row = block.title_row + k
    shifted_header_row = block.header_row + k
    for offset in range(2):
        source_row = shifted_title_row + offset
        target_row = insert_at + offset
        clone_row_dimension(ws, source_row, target_row)
        for col in range(1, ws.max_column + 1):
            copy_cell(ws.cell(source_row, col), ws.cell(target_row, col), include_value=True)
    new_header_merges: list[str] = []
    for merged in original_merges:
        if merged.min_row >= block.title_row and merged.max_row <= block.header_row:
            new_range = CellRange(
                min_col=merged.min_col,
                min_row=insert_at + (merged.min_row - block.title_row),
                max_col=merged.max_col,
                max_row=insert_at + (merged.max_row - block.title_row),
            )
            ws.merge_cells(str(new_range))
            new_header_merges.append(str(new_range))

    template_original_row = next(
        row for row in block.record_rows
        if "\n" not in str(ws.cell(row + k, block.title_col).value or "")
    )
    template_row = template_original_row + k
    precedent_rows = [row + k for row in block.record_rows]
    data_start = insert_at + 2
    row_height_precedents: list[dict[str, Any]] = []
    for index, record in enumerate(records):
        row = data_start + index
        clone_row_dimension(ws, template_row, row)
        for col in range(1, ws.max_column + 1):
            copy_cell(ws.cell(template_row, col), ws.cell(row, col), include_value=False)
            ws.cell(row, col).value = None
            ws.cell(row, col).hyperlink = None
            ws.cell(row, col).comment = None
        ws.cell(row, block.date_col).value = record.seen_date
        ws.cell(row, block.date_col).number_format = "yyyy/m/d"
        ws.cell(row, block.broker_col).value = record.broker
        ws.cell(row, block.title_col).value = record.title
        row_height_precedents.append(
            match_row_height_to_precedent(ws, block.title_col, row, precedent_rows)
        )

    date_values = [date_key(record.seen_date) for record in records]
    broker_values = [record.broker for record in records]
    new_merges: list[str] = []
    for start, end in contiguous_groups(date_values):
        if end > start:
            ref = f"{ws.cell(data_start + start, block.date_col).coordinate}:{ws.cell(data_start + end, block.date_col).coordinate}"
            ws.merge_cells(ref)
            new_merges.append(ref)
    for start, end in contiguous_groups(broker_values):
        if end > start:
            ref = f"{ws.cell(data_start + start, block.broker_col).coordinate}:{ws.cell(data_start + end, block.broker_col).coordinate}"
            ws.merge_cells(ref)
            new_merges.append(ref)

    candidates = []
    for index, record in enumerate(records):
        if is_board_resolution(record.title):
            row = data_start + index
            cell = ws.cell(row, block.title_col)
            cell.font = copy.copy(cell.font)
            cell.font = Font(
                name=cell.font.name,
                sz=cell.font.sz,
                b=True,
                i=cell.font.i,
                vertAlign=cell.font.vertAlign,
                underline=cell.font.underline,
                strike=cell.font.strike,
                color=copy.copy(cell.font.color),
                charset=cell.font.charset,
                family=cell.font.family,
                scheme=cell.font.scheme,
            )
            candidates.append({
                "broker": record.broker,
                "date": json_value(record.seen_date),
                "title": record.title,
                "sheet": ws.title,
                "cell": cell.coordinate,
                "source": record.source,
                "source_row": record.source_row,
            })

    output.parent.mkdir(parents=True, exist_ok=True)
    with warnings.catch_warnings(record=True) as save_warnings:
        warnings.simplefilter("always")
        wb.save(output)
    restore_printer_settings(master, output)
    package_report = verify_package(master, output)
    after_wb, reload_warnings = load_xlsx(output)
    after_ws = after_wb[ws.title]
    after_history = snapshot_sheet(after_ws, insert_at + k)
    after_other_sheets = workbook_sheet_snapshots(after_wb, exclude=ws.title)
    if before_history != after_history:
        raise EngineError(
            "Historical region verification failed after row shift: "
            + str(first_difference(before_history, after_history))
        )
    if before_other_sheets != after_other_sheets:
        raise EngineError("A non-target worksheet changed")
    if after_ws.sheet_state != before_sheet_state:
        raise EngineError("Target worksheet state changed")
    if after_ws.max_row != ws.max_row:
        raise EngineError("Reloaded workbook row count differs from saved workbook")
    after_merge_refs = {str(item) for item in after_ws.merged_cells.ranges}
    missing_header_merges = [ref for ref in new_header_merges if ref not in after_merge_refs]
    if missing_header_merges:
        raise EngineError(f"New header merge(s) were not preserved: {missing_header_merges}")

    effective_date = None
    effective_broker = None
    roundtrip_records = []
    for row in range(data_start, data_start + len(records)):
        date_value = after_ws.cell(row, block.date_col).value
        broker_value = after_ws.cell(row, block.broker_col).value
        if date_value not in (None, ""):
            effective_date = date_value
        if broker_value not in (None, ""):
            effective_broker = broker_value
        roundtrip_records.append((date_key(as_date(effective_date, after_wb.epoch)), effective_broker, str(after_ws.cell(row, block.title_col).value)))
    expected_records = [(date_key(record.seen_date), record.broker, record.title) for record in records]
    if roundtrip_records != expected_records:
        raise EngineError("New region does not reconcile to source records")
    for candidate in candidates:
        if after_ws[candidate["cell"]].font.bold is not True:
            raise EngineError(f"Board resolution title is not bold: {candidate['cell']}")
    for item in row_height_precedents:
        actual_height = after_ws.row_dimensions[item["row"]].height
        if actual_height != item["height"]:
            raise EngineError(
                f"New row height changed after reload at row {item['row']}: "
                f"expected {item['height']}, got {actual_height}"
            )

    report.update({
        "status": "complete",
        "output": str(output),
        "insert_count": k,
        "new_region": {
            "title_rows": [insert_at, insert_at + 1],
            "data_rows": [data_start, data_start + len(records) - 1] if records else None,
            "separator_after": insert_at + k if report["separator_row"] else None,
            "old_title_row_after": block.title_row + k,
        },
        "new_merges": new_merges,
        "new_header_merges": new_header_merges,
        "row_height_precedents": row_height_precedents,
        "board_resolution_candidates": candidates,
        "verification": {
            "history_unchanged": True,
            "other_sheets_unchanged": True,
            "source_records_reconciled": True,
            "new_row_heights_matched": True,
            "package": package_report,
            "warnings": load_warnings + [str(item.message) for item in save_warnings] + reload_warnings,
        },
    })
    return report


def inline_font_from_cell(cell: Cell, bold: bool) -> InlineFont:
    font = cell.font
    color = None
    if font.color is not None and font.color.type == "rgb":
        color = font.color.rgb
    return InlineFont(
        rFont=font.name,
        sz=font.sz,
        b=bold,
        i=font.i,
        u=font.u,
        strike=font.strike,
        color=color,
    )


def set_resolution_rich_text(cell: Cell, title: str, resolutions: list[str]) -> None:
    if not resolutions:
        raise EngineError("Resolution list cannot be empty")
    title_font = inline_font_from_cell(cell, True)
    body_font = inline_font_from_cell(cell, False)
    runs: list[Any] = [TextBlock(title_font, title)]
    for index, resolution in enumerate(resolutions, start=1):
        runs.append(TextBlock(body_font, f"\n{index}、{resolution}"))
    cell.value = CellRichText(runs)
    cell.font = copy.copy(cell.font)
    cell.font = Font(
        name=cell.font.name,
        sz=cell.font.sz,
        b=False,
        i=cell.font.i,
        color=copy.copy(cell.font.color),
        charset=cell.font.charset,
        family=cell.font.family,
        scheme=cell.font.scheme,
    )
    cell.alignment = copy.copy(cell.alignment)
    cell.alignment = Alignment(
        horizontal=cell.alignment.horizontal,
        vertical=cell.alignment.vertical,
        textRotation=cell.alignment.textRotation,
        wrapText=True,
        shrinkToFit=cell.alignment.shrinkToFit,
        indent=cell.alignment.indent,
        relativeIndent=cell.alignment.relativeIndent,
        justifyLastLine=cell.alignment.justifyLastLine,
        readingOrder=cell.alignment.readingOrder,
    )


def remove_stage2_allowed_differences(
    snapshot: dict[str, Any], targets: list[tuple[str, str]]
) -> dict[str, Any]:
    normalized = copy.deepcopy(snapshot)
    for sheet_name, coordinate in targets:
        if sheet_name not in normalized:
            continue
        match = re.fullmatch(r"([A-Z]+)(\d+)", coordinate)
        if not match:
            raise EngineError(f"Invalid target coordinate: {coordinate}")
        row = int(match.group(2))
        col = openpyxl.utils.column_index_from_string(match.group(1))
        sheet_snapshot = normalized[sheet_name]
        sheet_snapshot["cells"].pop(f"{row - 1}:{col}", None)
        sheet_snapshot["row_dimensions"].pop(str(row - 1), None)
    return normalized


def load_resolution_payload(path: Path) -> list[dict[str, Any]]:
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise EngineError(f"Cannot read resolution payload: {error}") from error
    entries = payload.get("resolutions") if isinstance(payload, dict) else None
    if not isinstance(entries, list) or not entries:
        raise EngineError("Resolution payload must contain a non-empty 'resolutions' list")
    required = {"broker", "title", "sheet", "cell", "pdf", "items"}
    for index, entry in enumerate(entries, start=1):
        if not isinstance(entry, dict) or not required.issubset(entry):
            raise EngineError(f"Resolution entry {index} is missing required fields")
        if not isinstance(entry["items"], list) or not entry["items"]:
            raise EngineError(f"Resolution entry {index} has no decision items")
        if any(not isinstance(item, str) or not item.strip() for item in entry["items"]):
            raise EngineError(f"Resolution entry {index} contains an empty decision item")
        if not Path(entry["pdf"]).is_file():
            raise EngineError(f"Source PDF does not exist: {entry['pdf']}")
    return entries


def stage2(
    input_workbook: Path,
    stage1_report_path: Path,
    resolutions_path: Path,
    output: Path,
) -> dict[str, Any]:
    if output.exists():
        raise EngineError(f"Refusing to overwrite existing output: {output}")
    try:
        stage1_report = json.loads(stage1_report_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise EngineError(f"Cannot read stage-one report: {error}") from error
    candidates = stage1_report.get("board_resolution_candidates")
    if stage1_report.get("status") != "complete" or not isinstance(candidates, list):
        raise EngineError("Stage-one report is incomplete or has no candidate list")
    entries = load_resolution_payload(resolutions_path)
    candidate_by_cell = {(item["sheet"], item["cell"]): item for item in candidates}
    entry_by_cell = {(item["sheet"], item["cell"]): item for item in entries}
    if len(entry_by_cell) != len(entries):
        raise EngineError("Resolution payload contains duplicate target cells")
    if set(candidate_by_cell) != set(entry_by_cell):
        missing = sorted(set(candidate_by_cell) - set(entry_by_cell))
        extra = sorted(set(entry_by_cell) - set(candidate_by_cell))
        raise EngineError(f"Resolution targets do not match stage-one candidates; missing={missing}, extra={extra}")

    wb, load_warnings = load_xlsx(input_workbook)
    before = workbook_sheet_snapshots(wb)
    targets = sorted(entry_by_cell)
    expected: dict[tuple[str, str], dict[str, Any]] = {}
    for key in targets:
        entry = entry_by_cell[key]
        candidate = candidate_by_cell[key]
        if entry["broker"] != candidate["broker"] or entry["title"] != candidate["title"]:
            raise EngineError(f"Resolution identity mismatch at {key[0]}!{key[1]}")
        if key[0] not in wb.sheetnames:
            raise EngineError(f"Target worksheet is missing: {key[0]}")
        cell = wb[key[0]][key[1]]
        if str(cell.value or "") != candidate["title"]:
            raise EngineError(f"Target title changed before stage two: {key[0]}!{key[1]}")
        set_resolution_rich_text(cell, candidate["title"], entry["items"])
        line_count = 1 + len(entry["items"])
        current_height = wb[key[0]].row_dimensions[cell.row].height or 15
        new_height = max(current_height, line_count * 15)
        wb[key[0]].row_dimensions[cell.row].height = new_height
        expected[key] = {
            "broker": entry["broker"],
            "title": entry["title"],
            "pdf": entry["pdf"],
            "items": entry["items"],
            "row_height": new_height,
        }

    output.parent.mkdir(parents=True, exist_ok=True)
    with warnings.catch_warnings(record=True) as save_warnings:
        warnings.simplefilter("always")
        wb.save(output)
    restore_printer_settings(input_workbook, output)
    package_report = verify_package(input_workbook, output)
    after_wb, reload_warnings = load_xlsx(output)
    after = workbook_sheet_snapshots(after_wb)
    before_normalized = remove_stage2_allowed_differences(before, targets)
    after_normalized = remove_stage2_allowed_differences(after, targets)
    if before_normalized != after_normalized:
        raise EngineError(
            "Workbook changed outside the approved stage-two cells: "
            + str(first_difference(before_normalized, after_normalized))
        )

    completed = []
    for key in targets:
        cell = after_wb[key[0]][key[1]]
        value = cell.value
        if not isinstance(value, CellRichText):
            raise EngineError(f"Rich text was not preserved at {key[0]}!{key[1]}")
        blocks = [item for item in value if isinstance(item, TextBlock)]
        if not blocks or blocks[0].font.b is not True:
            raise EngineError(f"Title is not bold at {key[0]}!{key[1]}")
        if any(item.font.b is True for item in blocks[1:]):
            raise EngineError(f"Decision body inherited bold at {key[0]}!{key[1]}")
        expected_text = expected[key]["title"] + "".join(
            f"\n{index}、{item}" for index, item in enumerate(expected[key]["items"], start=1)
        )
        if str(value) != expected_text:
            raise EngineError(f"Decision text mismatch at {key[0]}!{key[1]}")
        if cell.alignment.wrap_text is not True:
            raise EngineError(f"Text wrapping is not enabled at {key[0]}!{key[1]}")
        completed.append({
            "broker": expected[key]["broker"],
            "title": expected[key]["title"],
            "sheet": key[0],
            "cell": key[1],
            "pdf": expected[key]["pdf"],
            "decision_count": len(expected[key]["items"]),
            "row_height": expected[key]["row_height"],
        })

    return {
        "status": "complete",
        "input": str(input_workbook),
        "output": str(output),
        "completed_count": len(completed),
        "unresolved_count": 0,
        "completed": completed,
        "verification": {
            "outside_target_cells_unchanged": True,
            "title_only_bold": True,
            "body_not_bold": True,
            "all_items_numbered": True,
            "package": package_report,
            "warnings": load_warnings + [str(item.message) for item in save_warnings] + reload_warnings,
        },
    }


def make_selftest_files(root: Path) -> tuple[Path, list[Path]]:
    master = root / "master.xlsx"
    source_empty = root / "source_empty.xlsx"
    source_data = root / "source_data.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.column_dimensions["B"].width = 13
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 70
    header_fill = PatternFill("solid", fgColor="D9EAD3")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for title_row, header_row, data_row, stamp, broker, title in (
        (1, 2, 3, datetime(2026, 5, 28), "旧券商", "旧公告"),
        (6, 7, 8, datetime(2026, 8, 26), "甲券商", "昨日公告"),
    ):
        ws.cell(title_row, 2).value = "券商公告"
        ws.merge_cells(start_row=title_row, start_column=2, end_row=title_row, end_column=4)
        ws.cell(title_row, 2).font = Font(name="宋体", size=14, bold=True)
        for col, value in zip((2, 3, 4), ("见网日期", "券商名称", "公告题目")):
            ws.cell(header_row, col).value = value
            ws.cell(header_row, col).fill = header_fill
            ws.cell(header_row, col).border = border
            ws.cell(header_row, col).font = Font(name="宋体", size=12, bold=True)
        for col, value in zip((2, 3, 4), (stamp, broker, title)):
            ws.cell(data_row, col).value = value
            ws.cell(data_row, col).border = border
            ws.cell(data_row, col).font = Font(name="宋体", size=10)
        ws.cell(data_row, 2).number_format = "m/d/yy"
    ws["D8"].comment = Comment("history comment", "Tester")
    ws["D8"].hyperlink = "https://example.com/history"
    ws.row_dimensions[8].height = 22
    steps = wb.create_sheet("工作步骤")
    steps["A1"] = "保持不变"
    wb.save(master)
    for path, rows in (
        (source_empty, []),
        (source_data, [
            (datetime(2026, 8, 27), "乙券商", "普通公告"),
            (None, "乙券商", "乙券商第一届董事会第一次会议决议公告"),
            (None, "丙券商", "另一公告"),
        ]),
    ):
        source_wb = Workbook()
        source_ws = source_wb.active
        source_ws.append(["每日券商公告", None, None])
        source_ws.append(["挂网日期", "券商名称", "公告题目"])
        for row in rows:
            source_ws.append(list(row))
        source_wb.save(path)
    return master, [source_empty, source_data]


def selftest() -> dict[str, Any]:
    if not is_board_resolution("某公司第一届董事会第一次会议决议公告"):
        raise EngineError("Self-test standard board-resolution detection failed")
    if not is_board_resolution("关于第一届董事会第一次会议决议的公告"):
        raise EngineError("Self-test board-resolution '的公告' detection failed")
    if is_board_resolution("某公司第一届监事会第一次会议决议公告"):
        raise EngineError("Self-test supervisory-board exclusion failed")
    with tempfile.TemporaryDirectory(prefix="broker-daily-selftest-") as temp_dir:
        root = Path(temp_dir)
        master, sources = make_selftest_files(root)
        output = root / "stage1.xlsx"
        report = stage1(master, sources, output)
        if report["source_record_count"] != 3 or report["insert_count"] != 5:
            raise EngineError("Self-test record counts failed")
        if len(report["board_resolution_candidates"]) != 1:
            raise EngineError("Self-test board resolution detection failed")
        wb, _ = load_xlsx(output)
        candidate = report["board_resolution_candidates"][0]
        stage1_report_path = root / "stage1.json"
        stage1_report_path.write_text(
            json.dumps(report, ensure_ascii=False, default=json_value), encoding="utf-8"
        )
        pdf_source = root / "source.pdf"
        pdf_source.write_bytes(b"%PDF-1.4\n%%EOF\n")
        resolution_payload = root / "resolutions.json"
        resolution_payload.write_text(
            json.dumps({"resolutions": [{
                "broker": candidate["broker"],
                "title": candidate["title"],
                "sheet": candidate["sheet"],
                "cell": candidate["cell"],
                "pdf": str(pdf_source),
                "items": ["审议通过测试议案", "暂缓审议另一议案"],
            }]}, ensure_ascii=False),
            encoding="utf-8",
        )
        rich_output = root / "rich.xlsx"
        stage2(output, stage1_report_path, resolution_payload, rich_output)
        restored, _ = load_xlsx(rich_output)
        value = restored[candidate["sheet"]][candidate["cell"]].value
        if not isinstance(value, CellRichText):
            raise EngineError("Rich-text roundtrip did not preserve rich text")
        text_blocks = [item for item in value if isinstance(item, TextBlock)]
        if not text_blocks or text_blocks[0].font.b is not True:
            raise EngineError("Rich-text title is not bold")
        if any(item.font.b is True for item in text_blocks[1:]):
            raise EngineError("Rich-text resolution body incorrectly inherited bold")
        return {
            "status": "pass",
            "openpyxl": openpyxl.__version__,
            "stage1_history_check": True,
            "source_reconciliation": True,
            "rich_text_roundtrip": True,
        }


def write_json(path: Path | None, payload: dict[str, Any]) -> None:
    text = json.dumps(payload, ensure_ascii=False, indent=2, default=json_value)
    if path:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
    print(text)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("selftest")
    pre = subparsers.add_parser("preflight")
    pre.add_argument("--master", type=Path, required=True)
    pre.add_argument("--source", type=Path, action="append", required=True)
    pre.add_argument("--report", type=Path)
    run = subparsers.add_parser("stage1")
    run.add_argument("--master", type=Path, required=True)
    run.add_argument("--source", type=Path, action="append", required=True)
    run.add_argument("--output", type=Path, required=True)
    run.add_argument("--report", type=Path)
    finish = subparsers.add_parser("stage2")
    finish.add_argument("--input", type=Path, required=True)
    finish.add_argument("--stage1-report", type=Path, required=True)
    finish.add_argument("--resolutions", type=Path, required=True)
    finish.add_argument("--output", type=Path, required=True)
    finish.add_argument("--report", type=Path)
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    try:
        if args.command == "selftest":
            write_json(None, selftest())
        elif args.command == "preflight":
            write_json(args.report, preflight(args.master, args.source))
        elif args.command == "stage1":
            write_json(args.report, stage1(args.master, args.source, args.output))
        elif args.command == "stage2":
            write_json(
                args.report,
                stage2(args.input, args.stage1_report, args.resolutions, args.output),
            )
        return 0
    except EngineError as error:
        print(json.dumps({"status": "error", "error": str(error)}, ensure_ascii=False, indent=2))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
