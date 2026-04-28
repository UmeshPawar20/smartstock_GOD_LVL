"""
COPY / HANDOVER RUNNER (single command, no flags)

Goal
----
From a single raw DCR ZIP file containing 4 division XLSX inputs (Div28/30/35/42),
produce the validated COPY outputs for ALL divisions:
- SlotMAX multi-month CSV (baseline months + report month rows)
- Final Excel workbook with one sheet per month (DEC/JAN/FEB/MAR by default)
  + "New Mapping" column from Book3.xlsx
  + corrected "new prescriber" remark:
      total_rx(current_month) > 0 AND total_rx(previous N months) == 0

Usage
-----
Edit USER CONFIG below, then run:
  python run_copy_outputs_from_zip_ALL_DIVS.py
"""

from __future__ import annotations

import subprocess
import zipfile
from pathlib import Path
from typing import Dict, List

import pandas as pd


# ---------------------------
# USER CONFIG (edit these)
# ---------------------------
PROJECT_DIR = Path(__file__).resolve().parent

# Input ZIP from the business pack
ZIP_PATH = PROJECT_DIR / "DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div30.zip"

# Where to write outputs
OUTPUT_DIR = PROJECT_DIR / "_copy_outputs_from_zip"

# Report month:
# - "latest" uses latest month available in Date column
# - "YYYY-MM" fixed month, e.g. "2026-03"
REPORT_MONTH = "latest"

# Month sheets in output Excel
SHEET_MONTHS = ["DEC", "JAN", "FEB", "MAR"]

# New Mapping file (provided by user)
MAPPING_XLSX = PROJECT_DIR / "Book3.xlsx"
MAPPING_SHEET = "New Mapping"
MAPPING_KEY_COL = "Speciality By Practice"
MAPPING_VALUE_COL = "New Mapping"

# New prescriber lookback months (2 or 3)
LOOKBACK_MONTHS = 3

# Div scripts
DIV_SCRIPTS = {
    "28": PROJECT_DIR / "slotmax_brandwide_by_doctor_month_Div28.py",
    "30": PROJECT_DIR / "slotmax_brandwide_by_doctor_month_Div30.py",
    "35": PROJECT_DIR / "slotmax_brandwide_by_doctor_month_Div35.py",
    "42": PROJECT_DIR / "slotmax_brandwide_by_doctor_month_Div42.py",
}


# ---------------------------
# Helpers
# ---------------------------

def _norm_str(x: object) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    return str(x).strip()


def load_mapping(path: Path) -> Dict[str, str]:
    m = pd.read_excel(path, sheet_name=MAPPING_SHEET)
    cols = [c for c in m.columns if not str(c).lower().startswith("unnamed")]
    m = m[cols].copy()
    if MAPPING_KEY_COL not in m.columns or MAPPING_VALUE_COL not in m.columns:
        raise ValueError(f"Mapping sheet must contain columns: {MAPPING_KEY_COL!r}, {MAPPING_VALUE_COL!r}")
    m[MAPPING_KEY_COL] = m[MAPPING_KEY_COL].map(_norm_str)
    m[MAPPING_VALUE_COL] = m[MAPPING_VALUE_COL].map(_norm_str)
    m = m[(m[MAPPING_KEY_COL] != "") & (m[MAPPING_VALUE_COL] != "")]
    return dict(zip(m[MAPPING_KEY_COL], m[MAPPING_VALUE_COL]))


def drop_unnamed_columns(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[:, [c for c in df.columns if str(c).strip() != "" and not str(c).lower().startswith("unnamed:")]].copy()


def infer_brand_columns(df: pd.DataFrame, month_col: str = "MONTH") -> List[str]:
    cols = list(df.columns)
    if month_col not in cols:
        raise ValueError(f"MONTH column not found: {month_col}")

    def idx_of(name: str) -> int:
        for i, c in enumerate(cols):
            if str(c).strip().lower() == name.strip().lower():
                return i
        return -1

    i_month = idx_of("MONTH")
    i_sum = idx_of("Sum of last 3 months")
    i_var = idx_of("Variance >25%")
    end = i_sum if i_sum > i_month else i_var
    if end <= i_month:
        raise ValueError("Could not infer brand columns (expected Sum of last 3 months or Variance >25% after MONTH).")
    return [c for c in cols[i_month + 1 : end] if str(c).strip() != ""]


def compute_new_prescriber_remark(
    df: pd.DataFrame,
    brand_cols: List[str],
    lookback_months: int,
    sheet_months: List[str],
    doctor_col: str = "Account: Customer Code",
    month_col: str = "MONTH",
) -> pd.Series:
    out = pd.Series([""] * len(df), index=df.index, dtype="string")
    work = df[[doctor_col, month_col] + brand_cols].copy()
    for c in brand_cols:
        work[c] = pd.to_numeric(work[c], errors="coerce").fillna(0)
    work["_total"] = work[brand_cols].sum(axis=1)

    month_order = {m: i for i, m in enumerate([x.upper() for x in sheet_months])}
    work["_m_ord"] = work[month_col].astype(str).str.upper().map(lambda x: month_order.get(x, 9999))

    for _, sub in work.groupby(doctor_col, sort=False):
        sub = sub.sort_values("_m_ord", kind="mergesort")
        totals = sub["_total"].to_list()
        idxs = sub.index.to_list()
        for j in range(len(totals)):
            if totals[j] <= 0:
                continue
            start = max(0, j - lookback_months)
            prev = totals[start:j]
            if len(prev) == 0:
                continue
            if all(v == 0 for v in prev):
                out.loc[idxs[j]] = "new prescriber"

    return out


def extract_zip_inputs(zip_path: Path, extract_dir: Path) -> Dict[str, Path]:
    """
    Extract division XLSX files from zip into extract_dir.
    Returns dict div->xlsx_path.
    """
    extract_dir.mkdir(parents=True, exist_ok=True)
    want = {div: f"DCR_RAW_STANDARDIZED_4div_2026-03-01_2026-03-31_Div{div}.xlsx" for div in ["28", "30", "35", "42"]}

    with zipfile.ZipFile(zip_path) as z:
        names = set(z.namelist())
        out = {}
        for div, fname in want.items():
            if fname not in names:
                raise FileNotFoundError(f"ZIP missing expected file: {fname}")
            z.extract(fname, extract_dir)
            out[div] = extract_dir / fname
        return out


def run_slotmax(div: str, script_path: Path, input_xlsx: Path, out_csv: Path, report_month: str) -> None:
    cmd = [
        "python",
        str(script_path),
        "--input",
        str(input_xlsx),
        "--report-month",
        str(report_month),
        "--output",
        str(out_csv),
    ]
    r = subprocess.run(cmd, cwd=str(PROJECT_DIR))
    if r.returncode != 0:
        raise RuntimeError(f"SlotMAX failed for Div{div}: exit code {r.returncode}")


def write_monthly_excel_copy(
    in_csv: Path,
    out_xlsx: Path,
    mapping: Dict[str, str],
    sheet_months: List[str],
    lookback_months: int,
) -> None:
    df = pd.read_csv(in_csv, low_memory=False)
    df = drop_unnamed_columns(df)

    # New Mapping column
    src_col = "Assignment: Specialty By Practice"
    if src_col in df.columns:
        df["New Mapping"] = df[src_col].map(lambda x: mapping.get(_norm_str(x), ""))
    else:
        df["New Mapping"] = ""

    # Correct remark logic
    brand_cols = infer_brand_columns(df)
    df["Remark"] = compute_new_prescriber_remark(
        df,
        brand_cols=brand_cols,
        lookback_months=lookback_months,
        sheet_months=sheet_months,
    )

    out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        for m in sheet_months:
            sub = df[df["MONTH"].astype(str).str.upper() == m].copy()
            sub.to_excel(writer, sheet_name=m, index=False)


def main() -> None:
    if not ZIP_PATH.exists():
        raise FileNotFoundError(f"ZIP not found: {ZIP_PATH}")
    if not MAPPING_XLSX.exists():
        raise FileNotFoundError(f"Mapping XLSX not found: {MAPPING_XLSX}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    extract_dir = OUTPUT_DIR / "extracted_xlsx"
    intermediate_dir = OUTPUT_DIR / "intermediate_csv"
    final_dir = OUTPUT_DIR / "final_xlsx"
    intermediate_dir.mkdir(parents=True, exist_ok=True)
    final_dir.mkdir(parents=True, exist_ok=True)

    mapping = load_mapping(MAPPING_XLSX)
    div_inputs = extract_zip_inputs(ZIP_PATH, extract_dir)

    for div, input_xlsx in div_inputs.items():
        script = DIV_SCRIPTS[div]
        if not script.exists():
            raise FileNotFoundError(f"Division script not found: {script}")
        out_csv = intermediate_dir / f"out_slotmax_div{div}_from_zip_with_baseline_rows.csv"
        out_xlsx = final_dir / f"out_slotmax_div{div}_monthly_sheets_COPY.xlsx"

        print(f"Div{div}: SlotMAX -> {out_csv.name}", flush=True)
        run_slotmax(div, script, input_xlsx, out_csv, REPORT_MONTH)

        print(f"Div{div}: COPY Excel -> {out_xlsx.name}", flush=True)
        write_monthly_excel_copy(
            in_csv=out_csv,
            out_xlsx=out_xlsx,
            mapping=mapping,
            sheet_months=SHEET_MONTHS,
            lookback_months=LOOKBACK_MONTHS,
        )

    print(f"OK. Outputs written under: {OUTPUT_DIR}", flush=True)


if __name__ == "__main__":
    main()

