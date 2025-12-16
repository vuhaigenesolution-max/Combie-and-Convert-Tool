"""Combine Excel workbooks by run name and date.

This module builds a combined workbook per (runname, date) group. It is
designed to support the UI in Fontend/combie.py via ``combine_metadata``.
"""

from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd


@dataclass(frozen=True)
class SheetRule:
    """Configuration for how to merge a specific sheet.

    prefix_rows: number of top rows to preserve as-is from the first file.
    header_row: zero-based index of the row that holds the column headers.
    """

    prefix_rows: int
    header_row: int


DEFAULT_RULE = SheetRule(prefix_rows=0, header_row=0)
SHEET_RULES: dict[str, SheetRule] = {
    "Sample": SheetRule(prefix_rows=20, header_row=20),
    "Sample Import": SheetRule(prefix_rows=22, header_row=22),
}

FILENAME_REGEX = re.compile(r"^metadata_(?P<run>[A-Za-z0-9_-]+)_(?P<date>20\d{6})(?:_.*)?\.xlsx$", re.IGNORECASE)
# Regex bắt run name và ngày (YYYYMMDD) trong tên file, bỏ qua mọi suffix phía sau.


def _normalize_run(run: str) -> str:
    """Chuẩn hóa run name để tránh lặp prefix metadata_ và ký tự thừa."""
    run = re.sub(r"^metadata_", "", run, flags=re.IGNORECASE)
    return run.strip(" _-") or "run"


def combine_metadata(source_dir: str | Path, output_dir: str | Path) -> list[Path]:
    """Combine Excel files grouped by run name and date.

    Rules (per sheet):
    - Sheet ``Sample``: keep rows 1-20, row 21 is header, append data from row 22+.
    - Sheet ``Sample Import``: keep rows 1-22, row 23 is header, append data from row 24+.
    - Other sheets: use first row as header, append data from row 2+.

    Args:
        source_dir: Folder containing source Excel files.
        output_dir: Folder to write combined workbooks.

    Returns:
        List of created output file paths.
    """

    src_path = Path(source_dir)
    out_path = Path(output_dir)

    if not src_path.exists():
        raise FileNotFoundError(f"Source folder not found: {src_path}")
    out_path.mkdir(parents=True, exist_ok=True)

    files = _find_excel_files(src_path)
    if not files:
        raise FileNotFoundError("No Excel files found in source folder.")

    grouped: dict[tuple[str, str], list[Path]] = defaultdict(list)
    for file in files:
        key = _group_key(file)
        grouped[key].append(file)

    outputs: list[Path] = []
    for (runname, date_key), paths in grouped.items():
        combined = _combine_group(paths)  # Gộp từng sheet theo rule.
        norm_run = _normalize_run(runname)
        safe_runname = re.sub(r"[^A-Za-z0-9_-]+", "_", norm_run).strip("_") or "run"
        filename = f"metadata_{safe_runname}_{date_key}.xlsx"
        dest = out_path / filename
        _write_workbook(dest, combined)
        outputs.append(dest)

    return outputs


def combine_metadata_by_filename(source_dir: str | Path, output_dir: str | Path) -> list[Path]:
    """Combine Excel metadata files that follow the pattern metadata_<RUN>_<YYYYMMDD>[_<SUFFIX>].xlsx.

    Files are grouped by (RUN, YYYYMMDD) ignoring any suffix. Each group produces
    one output file named combined_metadata_<RUN>_<YYYYMMDD>.xlsx. Sheet merging
    rules are identical to ``combine_metadata``.
    """

    src_path = Path(source_dir)
    out_path = Path(output_dir)

    if not src_path.exists():
        raise FileNotFoundError(f"Source folder not found: {src_path}")
    out_path.mkdir(parents=True, exist_ok=True)

    grouped: dict[tuple[str, str], list[Path]] = defaultdict(list)
    for file in src_path.glob("*.xlsx"):
        if file.name.startswith("~$"):
            continue
        match = FILENAME_REGEX.match(file.name)
        if not match:
            continue
        run = _normalize_run(match.group("run"))
        date_key = match.group("date")
        grouped[(run, date_key)].append(file)

    if not grouped:
        raise FileNotFoundError("No matching metadata_<RUN>_<YYYYMMDD>.xlsx files found.")

    outputs: list[Path] = []
    for (run, date_key), paths in grouped.items():
        combined = _combine_group(paths)  # Dùng chung logic merge sheet theo rule.
        safe_run = re.sub(r"[^A-Za-z0-9_-]+", "_", run).strip("_") or "run"
        dest = out_path / f"metadata_{safe_run}_{date_key}.xlsx"
        _write_workbook(dest, combined)
        outputs.append(dest)

    return outputs


def _find_excel_files(folder: Path) -> list[Path]:
    patterns = ["*.xlsx", "*.xlsm", "*.xls"]
    files: list[Path] = []
    for pattern in patterns:
        files.extend(folder.glob(pattern))
    return [f for f in files if not f.name.startswith("~$")]


def _group_key(path: Path) -> tuple[str, str]:
    stem = path.stem
    date_match = re.search(r"(20\d{6})", stem)
    # Nếu tên không có YYYYMMDD, fallback dùng mtime của file.
    date_part = date_match.group(1) if date_match else datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y%m%d")
    run_part = _normalize_run(stem.replace(date_part, "").strip(" _-."))
    return run_part.lower(), date_part


def _combine_group(files: Iterable[Path]) -> dict[str, pd.DataFrame]:
    files = list(sorted(files))
    sheet_names = _collect_sheet_names(files)
    combined: dict[str, pd.DataFrame] = {}

    for sheet in sheet_names:
        rule = SHEET_RULES.get(sheet, DEFAULT_RULE)
        frames = _load_sheet_frames(files, sheet)
        if not frames:
            continue
        combined[sheet] = _merge_frames(frames, rule)  # Giữ prefix + header, nối data.

    return combined


def _collect_sheet_names(files: list[Path]) -> set[str]:
    names: set[str] = set()
    for file in files:
        try:
            xls = pd.ExcelFile(file)
            names.update(xls.sheet_names)
        except Exception:
            continue
    return names


def _load_sheet_frames(files: list[Path], sheet: str) -> list[pd.DataFrame]:
    frames: list[pd.DataFrame] = []
    for file in files:
        try:
            df = pd.read_excel(file, sheet_name=sheet, header=None, dtype=object)
            frames.append(df)
        except ValueError:
            # Sheet missing in this file; skip.
            continue
    return frames


def _merge_frames(frames: list[pd.DataFrame], rule: SheetRule) -> pd.DataFrame:
    base = frames[0]
    if base.empty:
        return base

    header_idx = rule.header_row
    prefix_rows = rule.prefix_rows
    if header_idx >= len(base.index):
        raise ValueError(f"Header row {header_idx + 1} missing in sheet")

    prefix = base.iloc[:prefix_rows].copy() if prefix_rows else pd.DataFrame()
    header_row = base.iloc[header_idx]
    base_cols = header_row.index
    data_start = header_idx + 1

    data_parts: list[pd.DataFrame] = []
    for frame in frames:
        if data_start > len(frame.index):
            continue
        part = frame.iloc[data_start:].copy()
        part = part.reindex(columns=base_cols)
        part = part.dropna(how="all")
        data_parts.append(part)

    merged_data = pd.concat(data_parts, ignore_index=True) if data_parts else pd.DataFrame(columns=base_cols)

    result = pd.concat(
        [
            prefix,
            pd.DataFrame([header_row]).reindex(columns=base_cols),
            merged_data,
        ],
        ignore_index=True,
    )

    return result


def _write_workbook(dest: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(dest, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False, header=False)
