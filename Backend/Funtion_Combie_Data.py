"""Các bước logic cho chức năng combine (bổ sung dần từng bước).

Logic 1: nhận đường dẫn nguồn (source) từ màn hình Combine.
"""

from __future__ import annotations

from pathlib import Path
import shutil
import re
import pandas as pd


def get_source_path(source_dir: str | Path) -> Path:
    """Logic 1: lấy và kiểm tra đường dẫn source do UI truyền xuống.

    - Nhận `source_dir` từ màn hình Combine (combie.py).
    - Trả về Path đã được chuẩn hóa.
    - Nếu không tồn tại, ném FileNotFoundError để UI hiển thị lỗi.
    """

    src = Path(source_dir)
    if not src.exists():
        raise FileNotFoundError(f"Source folder not found: {src}")
    return src


__all__ = ["get_source_path"]


# ------------------------------
# Logic 2: tìm và nhóm file metadata theo RUN + YYYYMMDD, bỏ qua hậu tố.
# ------------------------------

# Mẫu tên: metadata_<RUN>_<YYYYMMDD>[_suffix].xlsx
FILENAME_REGEX = re.compile(r"^metadata_(?P<run>[A-Za-z0-9_-]+)_(?P<date>20\d{6})(?:_.*)?\.xlsx$", re.IGNORECASE)


def find_metadata_files(source_dir: Path) -> list[Path]:
    """Quét thư mục, lấy danh sách file metadata hợp lệ; báo lỗi nếu tên sai.

    - Hợp lệ: metadata_<RUN>_<YYYYMMDD>[_suffix].xlsx
    - Bỏ qua file tạm bắt đầu bằng ~$.
    - Nếu có file sai định dạng, ném ValueError kèm danh sách tên sai.
    """

    files: list[Path] = []
    invalid: list[str] = []
    for path in source_dir.glob("*.xlsx"):
        if path.name.startswith("~$"):
            continue  # file tạm của Excel
        if FILENAME_REGEX.match(path.name):
            files.append(path)
        else:
            invalid.append(path.name)

    if invalid:
        raise ValueError(
            "Các file không đúng định dạng metadata_<RUN>_<YYYYMMDD>[_suffix].xlsx: "
            + ", ".join(invalid)
        )

    return files


def group_by_run_date(files: list[Path]) -> dict[tuple[str, str], list[Path]]:
    """Nhóm các file theo (run, date) bất kể hậu tố phía sau.

    - run lấy từ nhóm (?P<run>)
    - date lấy từ nhóm (?P<date>) dạng YYYYMMDD
    - hậu tố (nếu có) bị bỏ qua, miễn là cùng run + date thì chung nhóm
    """

    groups: dict[tuple[str, str], list[Path]] = {}
    for f in files:
        m = FILENAME_REGEX.match(f.name)
        if not m:
            continue
        key = (m.group("run"), m.group("date"))
        groups.setdefault(key, []).append(f)
    return groups


__all__.extend([
    "find_metadata_files",
    "group_by_run_date",
])


# ------------------------------
# Logic 3: tạo tên file đầu ra và thư mục đích.
# ------------------------------


def get_output_path(output_dir: str | Path) -> Path:
    """Kiểm tra/thạo thư mục đích, trả về Path."""

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    return out


def build_output_filename(run: str, date: str) -> str:
    """Ghép tên file đầu ra: metadata_<run>_<YYYYMMDD>.xlsx"""

    return f"metadata_{run}_{date}.xlsx"


def plan_outputs(source_dir: str | Path, output_dir: str | Path) -> list[Path]:
    """Từ thư mục nguồn, xác định các nhóm (run, date) và tên file output.

    - Đọc source_dir, lấy file metadata hợp lệ.
    - Nhóm theo (run, date) bất kể hậu tố.
    - Tạo danh sách đường dẫn output tương ứng dưới output_dir.
    - Chưa gộp nội dung, chỉ trả về kế hoạch tên file đầu ra.
    """

    src = get_source_path(source_dir)
    out = get_output_path(output_dir)
    files = find_metadata_files(src)
    groups = group_by_run_date(files)

    outputs: list[Path] = []
    for (run, date), _paths in groups.items():
        outputs.append(out / build_output_filename(run, date))
    return outputs


__all__.extend([
    "get_output_path",
    "build_output_filename",
    "plan_outputs",
])


# ------------------------------
# Logic 4: gộp file theo nhóm run/date và lưu vào output.
# ------------------------------

# Rule sheet: Sample giữ dòng 1-20, header dòng 21, data từ dòng 22; Sample Import giữ dòng 1-22, header dòng 23, data từ dòng 24; sheet khác: header dòng 1, data từ dòng 2.
SHEET_RULES = {
    "Sample": {"prefix": 20, "header": 20},
    "Sample Import": {"prefix": 22, "header": 22},
}


def _collect_sheet_names(files: list[Path]) -> set[str]:
    names: set[str] = set()
    for f in files:
        try:
            xl = pd.ExcelFile(f)
            names.update(xl.sheet_names)
        except Exception:
            continue
    return names


def _load_sheet_frames(files: list[Path], sheet: str) -> list[pd.DataFrame]:
    frames: list[pd.DataFrame] = []
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name=sheet, header=None, dtype=object)
            frames.append(df)
        except ValueError:
            continue
    return frames


def _merge_frames(frames: list[pd.DataFrame], sheet: str) -> pd.DataFrame:
    rule = SHEET_RULES.get(sheet, {"prefix": 0, "header": 0})
    prefix_rows = rule["prefix"]
    header_idx = rule["header"]

    base = frames[0]
    if header_idx >= len(base.index):
        raise ValueError(f"Sheet {sheet}: thiếu dòng header (vị trí {header_idx + 1})")

    prefix = base.iloc[:prefix_rows].copy() if prefix_rows else pd.DataFrame()
    header_row = base.iloc[header_idx]
    cols = header_row.index
    data_start = header_idx + 1

    parts: list[pd.DataFrame] = []
    for fr in frames:
        if data_start > len(fr.index):
            continue
        part = fr.iloc[data_start:].copy()
        part = part.reindex(columns=cols)
        part = part.dropna(how="all")
        parts.append(part)

    merged_data = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame(columns=cols)

    if sheet == "Sample":
        # Bỏ dòng trống ở cả expNum và LABCODE
        header_map = {str(v).strip(): c for c, v in header_row.items() if str(v).strip()}
        exp_col = header_map.get("expNum")
        lab_col = header_map.get("LABCODE")
        if exp_col in merged_data.columns and lab_col in merged_data.columns:
            def _non_empty(series: pd.Series) -> pd.Series:
                return series.notna() & (series.astype(str).str.strip() != "")
            mask = _non_empty(merged_data[exp_col]) & _non_empty(merged_data[lab_col])
            merged_data = merged_data[mask].reset_index(drop=True)

    if sheet == "Sample Import":
        # Bỏ dòng có Sample_ID trống hoặc chỉ là "-"
        header_map = {str(v).strip(): c for c, v in header_row.items() if str(v).strip()}
        sample_col = header_map.get("Sample_ID")
        if sample_col in merged_data.columns:
            def _valid(series: pd.Series) -> pd.Series:
                s = series.astype(str).str.strip()
                return series.notna() & (s != "") & (s != "-")
            merged_data = merged_data[_valid(merged_data[sample_col])].reset_index(drop=True)

    result = pd.concat([prefix, pd.DataFrame([header_row]).reindex(columns=cols), merged_data], ignore_index=True)
    return result


def _write_workbook(dest: Path, sheets: dict[str, pd.DataFrame]) -> None:
    with pd.ExcelWriter(dest, engine="openpyxl") as writer:
        for name, df in sheets.items():
            df.to_excel(writer, sheet_name=name, index=False, header=False)


def combine_metadata_by_filename(source_dir: str | Path, output_dir: str | Path) -> list[Path]:
    """Gộp các file metadata cùng run/date và lưu vào output đã chọn.

    - Nhóm theo (RUN, YYYYMMDD) bỏ qua hậu tố.
    - Với mỗi nhóm: gộp sheet theo rule (Sample, Sample Import, sheet khác append theo header dòng đầu).
    - Xuất 1 file duy nhất: metadata_<RUN>_<YYYYMMDD>.xlsx trong thư mục output do UI chọn.
    """

    src = get_source_path(source_dir)
    out = get_output_path(output_dir)
    files = find_metadata_files(src)
    groups = group_by_run_date(files)

    if not groups:
        raise FileNotFoundError("Không tìm thấy file metadata hợp lệ trong thư mục nguồn.")

    outputs: list[Path] = []
    for (run, date), paths in groups.items():
        sheet_names = _collect_sheet_names(paths)
        combined: dict[str, pd.DataFrame] = {}
        for sheet in sheet_names:
            frames = _load_sheet_frames(paths, sheet)
            if not frames:
                continue
            combined[sheet] = _merge_frames(frames, sheet)

        dest = out / build_output_filename(run, date)
        _write_workbook(dest, combined)
        outputs.append(dest)

    return outputs


__all__.extend([
    "combine_metadata_by_filename",
])

