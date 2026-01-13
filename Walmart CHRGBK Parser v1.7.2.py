
#!/usr/bin/env python3
"""
Walmart_CB_Parser_v1.7.2.py
- Reads raw Check_*.xlsx files directly (preserves order)
- Outputs single sheet "Walmart_Deductions_All"
- Inserts "Check No" at column B derived from source filename (e.g., 001256261), preserving leading zeros
- Inserts "Internal Invoice Date" at column C and populates per 3-check nearest-invoice-below logic
- Removes the helper column that previously sat at O (internal row index), and also removes the raw file column
- Applies short date format
"""
import argparse, os, sys, re
from pathlib import Path
import pandas as pd
import numpy as np

def to_num(x):
    if pd.isna(x): return np.nan
    s = str(x).strip().replace("$","").replace(",","")
    if s.startswith("(") and s.endswith(")"):
        s = "-" + s[1:-1]
    try:
        return float(s)
    except Exception:
        return np.nan

def find_col(df, names):
    low = {str(c).strip().lower(): c for c in df.columns}
    for n in names:
        if n in low:
            return low[n]
    return None

def normalize_str(x):
    if pd.isna(x): return None
    return str(x).strip()

def read_check_files(input_dir: Path) -> pd.DataFrame:
    frames = []
    for p in sorted(input_dir.glob("Check_*.xlsx")):
        try:
            xls = pd.ExcelFile(p)
            first = xls.sheet_names[0]
            df = xls.parse(first)
            df["_file"] = p.name
            df["_row_in_file"] = np.arange(len(df))
            # Extract check number string (preserve leading zeros)
            m = re.search(r"Check_(\d+)\.xlsx$", p.name, flags=re.IGNORECASE)
            check_no = m.group(1) if m else ""
            df["Check No"] = check_no
            frames.append(df)
        except Exception as e:
            print(f"WARNING: failed to read {p}: {e}", file=sys.stderr)
    if frames:
        all_df = pd.concat(frames, ignore_index=True, sort=False)
        all_df = all_df.sort_values(by=["_file", "_row_in_file"], kind="stable").reset_index(drop=True)
    else:
        all_df = pd.DataFrame()
    return all_df

def fill_internal_invoice_dates(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        df["Internal Invoice Date"] = pd.NaT
        return df

    inv_date_col = find_col(df, ["invoice date"])
    amt_col      = find_col(df, ["amount paid($)","amount paid ($)","amount paid"])
    inv_num_col  = find_col(df, ["invoice number"])
    po_col       = find_col(df, ["po number","po #","po"])
    store_col    = find_col(df, ["store number","store #","store"])
    dc_col       = find_col(df, ["dc number","dc #","dc"])
    div_col      = find_col(df, ["division"])

    # Ensure "Internal Invoice Date" column exists
    if "Internal Invoice Date" not in df.columns:
        df["Internal Invoice Date"] = pd.NaT

    if inv_date_col is not None:
        df[inv_date_col] = pd.to_datetime(df[inv_date_col], errors="coerce")

    if amt_col is None:
        return df

    amt_num = df[amt_col].apply(to_num)
    inv_num_vals = df[inv_num_col].apply(normalize_str) if inv_num_col else pd.Series([None]*len(df))
    po_vals      = df[po_col].apply(normalize_str) if po_col else pd.Series([None]*len(df))
    store_vals   = df[store_col].apply(normalize_str) if store_col else pd.Series([None]*len(df))
    dc_vals      = df[dc_col].apply(normalize_str) if dc_col else pd.Series([None]*len(df))
    div_vals     = df[div_col].apply(normalize_str) if div_col else pd.Series([None]*len(df))

    files = df["_file"].tolist() if "_file" in df.columns else [None]*len(df)
    for i in range(len(df)):
        if pd.isna(amt_num.iloc[i]) or amt_num.iloc[i] >= 0:
            continue  # only deductions
        file_i = files[i]
        key_invoice = inv_num_vals.iloc[i]
        key_po      = po_vals.iloc[i]
        s_val, d_val, v_val = store_vals.iloc[i], dc_vals.iloc[i], div_vals.iloc[i]

        match_date = pd.NaT
        for j in range(i+1, len(df)):
            if files[j] != file_i:
                continue
            if pd.isna(amt_num.iloc[j]) or amt_num.iloc[j] <= 0:
                continue
            same_store = (s_val is None or store_vals.iloc[j] is None) or (store_vals.iloc[j] == s_val)
            same_dc    = (d_val is None or dc_vals.iloc[j] is None) or (dc_vals.iloc[j] == d_val)
            same_div   = (v_val is None or div_vals.iloc[j] is None) or (div_vals.iloc[j] == v_val)
            if key_invoice:
                same_key = (inv_num_vals.iloc[j] == key_invoice)
            else:
                same_key = (po_vals.iloc[j] == key_po) if key_po else False
            if same_key and same_store and same_dc and same_div:
                match_date = df[inv_date_col].iloc[j] if inv_date_col else pd.NaT
                break
        df.at[i, "Internal Invoice Date"] = match_date

    # --- Override rule: Deduction code 0780 Transportation related billing ---
    try:
        ded_code_col = find_col(df, ["deduction code"])
    except Exception:
        ded_code_col = None
    if ded_code_col is not None and inv_date_col is not None:
        ded_series = df[ded_code_col].astype(str).str.lower()
        mask_0780 = ded_series.str.contains(r"\b0780\b") | ded_series.str.contains("transportation related billing")
        if mask_0780.any():
            df.loc[mask_0780, "Internal Invoice Date"] = df.loc[mask_0780, inv_date_col]
    # --- end override ---

    return df

def final_order(df: pd.DataFrame) -> pd.DataFrame:
    # Ensure Check No at B and Internal Invoice Date at C
    cols = list(df.columns)
    # 1) Move/insert "Check No" to index 1
    if "Check No" not in cols:
        df.insert(1, "Check No", "")
        cols = list(df.columns)
    if cols.index("Check No") != 1:
        cols.insert(1, cols.pop(cols.index("Check No")))
    # 2) Ensure "Internal Invoice Date" at index 2
    if "Internal Invoice Date" not in cols:
        df.insert(2, "Internal Invoice Date", pd.NaT)
        cols = list(df.columns)
    if cols.index("Internal Invoice Date") != 2:
        cols.insert(2, cols.pop(cols.index("Internal Invoice Date")))
    # 3) Drop column O (15th) if it exists in this frame
    if len(cols) >= 15:
        # Determine which name sits at O (index 14)
        col_O = cols[14]
        cols.remove(col_O)
        df = df[cols]
    else:
        df = df[cols]
    # 4) Drop helper columns if present
    for helper in ["_row_in_file", "_file"]:
        if helper in df.columns:
            df = df.drop(columns=[helper])
    return df


def clean_rows_postparse(df: pd.DataFrame) -> pd.DataFrame:
    """
    After parsing:
    - Drop any rows that contain '|' in any cell
    - Drop rows where the "Invoice Number" column is blank/NaN.
      If "Invoice Number" is not found, fall back to column D (index 3) if present.
    """
    if df is None or df.empty:
        return df

    # Remove rows that contain '|' anywhere (treat NaN and datetimes safely)
    def contains_pipe(series):
        if str(series.dtype).startswith('datetime64'):
            return pd.Series([False] * len(series))
        return series.astype(str).str.contains(r"\|", na=False)

    mask_pipe = df.apply(contains_pipe, axis=0)
    rows_with_pipe = mask_pipe.any(axis=1)
    df = df.loc[~rows_with_pipe].copy()

    # Determine the Invoice Number column by name, with fallback to 4th column
    inv_name_candidates = [c for c in df.columns if str(c).strip().lower() == "invoice number"]
    if inv_name_candidates:
        inv_col = inv_name_candidates[0]
    elif len(df.columns) >= 4:
        inv_col = df.columns[3]
    else:
        # No viable column to validate; return as-is
        return df

    original_series = df[inv_col]
    mask_has_val = original_series.notna() & original_series.astype(str).str.strip().ne("")
    df = df.loc[mask_has_val].copy()

    return df


    # Remove rows that contain '|' anywhere (treat NaN and datetimes safely)
    def contains_pipe(series):
        if str(series.dtype).startswith('datetime64'):
            return pd.Series([False] * len(series))
        return series.astype(str).str.contains(r"\|", na=False)

    mask_pipe = df.apply(contains_pipe, axis=0)
    rows_with_pipe = mask_pipe.any(axis=1)
    df = df.loc[~rows_with_pipe].copy()

    # Require an Invoice value in column D (4th column)
    if len(df.columns) >= 4:
        inv_col = df.columns[3]
        original_series = df[inv_col]
        mask_has_val = original_series.notna() & original_series.astype(str).str.strip().ne("")
        df = df.loc[mask_has_val].copy()

    return df


    # Remove rows that contain '|' anywhere (treat NaN and datetimes safely)
    def contains_pipe(series):
        if str(series.dtype).startswith('datetime64'):
            return pd.Series([False] * len(series))
        return series.astype(str).str.contains(r"\|", na=False)

    mask_pipe = df.apply(contains_pipe, axis=0)
    rows_with_pipe = mask_pipe.any(axis=1)
    df = df.loc[~rows_with_pipe].copy()

    # Require an Invoice value in column D (4th column)
    if len(df.columns) >= 4:
        inv_col = df.columns[3]
        df[inv_col] = df[inv_col].astype(str).str.strip()
        df = df.loc[df[inv_col].notna() & (df[inv_col] != "")].copy()

    return df




def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Directory with Check_*.xlsx files")
    ap.add_argument("--output", required=True, help="Final output .xlsx path")
    args = ap.parse_args()

    input_dir = Path(args.input)
    df = read_check_files(input_dir)
    df = fill_internal_invoice_dates(df)
    df = final_order(df)
    df = clean_rows_postparse(df)

    with pd.ExcelWriter(args.output, engine="xlsxwriter", datetime_format="mm/dd/yyyy", date_format="mm/dd/yyyy") as writer:
        df.to_excel(writer, sheet_name="Deductions_All", index=False)
        wb = writer.book
        ws = writer.sheets["Deductions_All"]
        date_fmt = wb.add_format({"num_format": "mm/dd/yyyy"})
        for idx, col in enumerate(df.columns):
            if ("date" in str(col).lower()) or pd.api.types.is_datetime64_any_dtype(df[col]):
                ws.set_column(idx, idx, 12, date_fmt)
        if len(df.columns) > 0:
            ws.set_column(0, len(df.columns)-1, 14)
        # Add Excel table formatting across the written range
        nrows, ncols = df.shape
        last_row = nrows  # include header row
        last_col = ncols - 1
        if ncols > 0:
            ws.add_table(0, 0, last_row, last_col, {
                "name": "WalmartDeductionsTable",
                "columns": [{"header": str(col)} for col in df.columns]
            })

    print(f"Single-sheet workbook written to: {args.output}")

if __name__ == "__main__":
    main()
