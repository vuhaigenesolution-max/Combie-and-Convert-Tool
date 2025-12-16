"""Utilities to combine Excel metadata files by run and date.

Flow:
- Scan source folder for Excel files matching pattern metadata_<RUNNAME>_<YYYYMMDD>[_<SUFFIX>].xlsx.
- Group files by (RUNNAME, date), ignoring suffix.
- For each group: read sheet 'Sample', keep rows 1-20 as metadata, row 21 as header (once),
  append data rows (22+) from all files, and restrict columns to DESIRED_COLUMNS in that order.
- Write combined file as metadata_<RUNNAME>_<YYYYMMDD>.xlsx with sheet 'Sample'.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet


_PATTERN = re.compile(r"^metadata_(?P<run>.+?)_(?P<date>\d{8})(?:_.+)?\.xlsx$", re.IGNORECASE)
DESIRED_COLUMNS = [
	"expNum",
	"sampleOrder",
	"LABCODE",
	"Library By",
	"Library Date",
	"Species",
	"i7 index",
	"i5 index",
	"LibraryConc",
	"LibraryAmp",
	"Library Protocol",
	"LRMtemplate",
	"passedQC",
	"LANE",
	"Primers",
]

HEADER_ROW_INDEX = 23  # header line for SampleImport
DATA_START_ROW = HEADER_ROW_INDEX + 1  # data starts at row 24

OUTPUT_COLUMNS = [
	"Sample_ID",
	"Sample_Name",
	"Sample_Plate",
	"Sample_Well",
	"Index_Plate_Well",
	"I7_Index_ID",
	"index",
	"I5_Index_ID",
	"index2",
	"Sample_Project",
	"Description",
]


@dataclass
class CombineResult:
	output_path: Path
	run_name: str
	run_date: str
	source_files: List[Path]
	errors: List[str]


def _safe_mkdir(path: Path) -> None:
	path.mkdir(parents=True, exist_ok=True)


def _parse_groups(source_folder: Path) -> dict[tuple[str, str], list[Path]]:
	"""Group files by (run, date) based on filename pattern."""
	groups: dict[tuple[str, str], list[Path]] = {}
	for path in source_folder.glob("*.xlsx"):
		match = _PATTERN.match(path.name)
		if not match:
			continue
		key = (match.group("run"), match.group("date"))
		groups.setdefault(key, []).append(path)
	return groups


def _read_sample_sheet(path: Path) -> Optional[pd.DataFrame]:
	"""Safely read sheet 'Sample' without headers (managed manually)."""
	try:
		return pd.read_excel(path, sheet_name="Sample", header=None, engine="openpyxl")
	except Exception as exc:  # pragma: no cover - defensive read
		print(f"[WARN] Failed to read {path.name}: {exc}")
		return None


def _build_header_index(ws: Worksheet, header_row: int) -> Dict[str, int]:
	"""Map header value to column index for a given row."""
	headers: Dict[str, int] = {}
	for idx, cell in enumerate(ws[header_row], start=1):
		if cell.value is None:
			continue
		key = str(cell.value).strip()
		if key:
			headers[key] = idx
	return headers


def _get_cell(ws: Worksheet, row: int, col: int):
	return ws.cell(row=row, column=col).value


def create_sample_import(workbook_path: str | Path, sheet_name: str = "SampleImport") -> None:
	"""Create/replace SampleImport sheet inside the same workbook, preserving Sample sheet."""
	wb = load_workbook(workbook_path)
	if "Sample" not in wb.sheetnames:
		raise ValueError("Sheet 'Sample' not found")

	sample_ws = wb["Sample"]

	# Replace target sheet if exists.
	if sheet_name in wb.sheetnames:
		del wb[sheet_name]
	ws = wb.create_sheet(sheet_name)

	# Copy first 22 rows (metadata area) from Sample into SampleImport.
	max_copy_row = min(HEADER_ROW_INDEX - 1, sample_ws.max_row)
	max_copy_col = sample_ws.max_column
	for r in range(1, max_copy_row + 1):
		for c in range(1, max_copy_col + 1):
			ws.cell(row=r, column=c, value=_get_cell(sample_ws, r, c))

	# Build header map from Sample (assume header at row 21 in Sample).
	sample_header_row = 21
	header_index = _build_header_index(sample_ws, sample_header_row)

	# Set B23 = Sample!C14 as requested.
	ws.cell(row=HEADER_ROW_INDEX, column=2, value=_get_cell(sample_ws, 14, 3))

	# Write column headers at row 23.
	for col_idx, name in enumerate(OUTPUT_COLUMNS, start=1):
		ws.cell(row=HEADER_ROW_INDEX, column=col_idx, value=name)

	# Data rows from Sample (row 22 onward in Sample).
	out_row = DATA_START_ROW
	max_row = sample_ws.max_row
	for r in range(sample_header_row + 1, max_row + 1):
		if all(sample_ws.cell(row=r, column=c).value in (None, "") for c in range(1, sample_ws.max_column + 1)):
			continue

		def val(col_name: str):
			idx = header_index.get(col_name)
			return _get_cell(sample_ws, r, idx) if idx else None

		sample_order = val("sampleOrder")
		labcode = val("LABCODE")
		exp_num = val("expNum")
		i7 = val("i7 index")
		i5 = val("i5 index")
		idx1 = val("index")
		idx2 = val("index2")
		lane = val("LANE")

		def join_id(a, b):
			parts = [str(x).strip() for x in (a, b) if x not in (None, "")]
			return " - ".join(parts) if parts else None

		sample_id = join_id(sample_order, labcode)

		row_values = [
			sample_id,
			sample_id,
			exp_num,
			None,
			None,
			i7,
			idx1,
			i5,
			idx2,
			lane,
			None,
		]

		for c_idx, value in enumerate(row_values, start=1):
			ws.cell(row=out_row, column=c_idx, value=value)
		out_row += 1

	wb.save(workbook_path)


def combine_metadata(source_folder: str | Path, output_folder: str | Path) -> list[CombineResult]:
	"""
	Combine metadata Excel files by Run Name and Run Date.

	- Input files pattern: metadata_<RUNNAME>_<YYYYMMDD>[_<SUFFIX>].xlsx
	- Grouped by (RUNNAME, YYYYMMDD), ignoring suffix.
	- Sheet 'Sample' only: rows 1-20 metadata (keep from first file only), row 21 header (once), rows 22+ data (append).
	- Writes combined_metadata_<RUNNAME>_<YYYYMMDD>.xlsx per group.

	Returns list of CombineResult per written file.
	"""

	src = Path(source_folder)
	out = Path(output_folder)
	_safe_mkdir(out)

	groups = _parse_groups(src)
	results: list[CombineResult] = []

	if not groups:
		print(f"[INFO] No matching files in {src}")
		return results

	for (run_name, run_date), files in groups.items():
		errors: list[str] = []
		combined_data_frames: list[pd.DataFrame] = []
		metadata_rows: Optional[pd.DataFrame] = None
		header_row: Optional[pd.Series] = None

		for idx, path in enumerate(sorted(files)):
			df = _read_sample_sheet(path)
			if df is None:
				errors.append(f"Read failed: {path.name}")
				continue

			if df.shape[0] < 21:
				errors.append(f"Sheet too short (needs >=21 rows): {path.name}")
				continue

			if idx == 0:
				metadata_rows = df.iloc[:20].copy()
				header_row = df.iloc[20].copy()

				# Determine positions of desired headers; keep None where missing to fill with nulls.
				header_values = [str(v).strip() for v in header_row.tolist()]
				keep_positions: list[Optional[int]] = []
				for col_name in DESIRED_COLUMNS:
					keep_positions.append(header_values.index(col_name) if col_name in header_values else None)
				if all(pos is None for pos in keep_positions):
					errors.append(f"No desired columns found in header: {path.name}")
					continue
			# If we did not find keep_positions, skip this file.
			if metadata_rows is None or header_row is None:
				continue

			data_rows = df.iloc[21:]
			# Build aligned data frame with fixed columns; missing columns filled with NA.
			aligned_cols = []
			for pos in keep_positions:
				if pos is None:
					aligned_cols.append(pd.Series([pd.NA] * len(data_rows), index=data_rows.index))
				else:
					aligned_cols.append(data_rows.iloc[:, pos])
			aligned_df = pd.concat(aligned_cols, axis=1)
			aligned_df.columns = DESIRED_COLUMNS
			combined_data_frames.append(aligned_df)

		if not combined_data_frames or metadata_rows is None or header_row is None:
			print(f"[WARN] No valid data for group {run_name} {run_date}; skipped")
			continue

		combined_data = pd.concat(combined_data_frames, ignore_index=True)

		# Build final DataFrame: metadata (trimmed/fill null), header (fixed), then data (aligned).
		meta_cols = []
		for pos in keep_positions:
			if pos is None:
				meta_cols.append(pd.Series([pd.NA] * len(metadata_rows), index=metadata_rows.index))
			else:
				meta_cols.append(metadata_rows.iloc[:, pos])
		metadata_trimmed = pd.concat(meta_cols, axis=1)
		metadata_trimmed.columns = DESIRED_COLUMNS

		header_as_df = pd.DataFrame([DESIRED_COLUMNS])

		final_df = pd.concat([metadata_trimmed, header_as_df, combined_data], ignore_index=True)

		output_name = f"metadata_{run_name}_{run_date}.xlsx"
		output_path = out / output_name

		try:
			final_df.to_excel(output_path, index=False, header=False, sheet_name="Sample", engine="openpyxl")
		except Exception as exc:  # pragma: no cover - defensive write
			print(f"[ERROR] Failed to write {output_name}: {exc}")
			errors.append(f"Write failed: {exc}")
			continue

		# Append SampleImport sheet after Sample is written.
		try:
			create_sample_import(output_path, sheet_name="SampleImport")
		except Exception as exc:  # pragma: no cover - defensive sheet creation
			msg = f"Create SampleImport failed for {output_name}: {exc}"
			print(f"[WARN] {msg}")
			errors.append(msg)

		results.append(CombineResult(output_path=output_path, run_name=run_name, run_date=run_date, source_files=sorted(files), errors=errors))
		print(f"[INFO] Wrote {output_path} (Sample + SampleImport)")

	return results


__all__ = ["combine_metadata", "CombineResult"]
