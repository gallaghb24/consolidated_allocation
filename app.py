"""allocation_merger_app.py – v3
Keeps full store‑detail columns for every store, even when that store first
appears in a *later* file.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

# ──────────────────────────────────────────────────────────────────────────────
#  Dependencies check
# ──────────────────────────────────────────────────────────────────────────────
try:
    import openpyxl  # noqa: F401
except ImportError:
    st.error("❌ `openpyxl` is missing. Add it to *requirements.txt* or run `pip install openpyxl`.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
#  Constants & helpers
# ──────────────────────────────────────────────────────────────────────────────
KEY_COLS = [
    "Store Number",
    "Store Name",
    "Address Line 1",
    "Address Line 2",
    "City or Town",
    "County",
    "Country",
    "Post Code",
    "Region / Area",
    "Location Type",
    "Trading Format",
]
LABELS = [
    "Brief Description",
    "Total (inc Overs)",
    "Total Allocations",
    "Overs",
]
LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1  # 1‑based Excel column where labels go (K)
ITEM_START_XL = LABEL_COL_XL + 1                      # first item column (L)


def extract_data_and_meta(file):
    """Return (df, meta_dict) for a single allocation export."""
    df = pd.read_excel(file, header=6, engine="openpyxl")
    df["Store Number"] = df["Store Number"].astype(str)

    raw = pd.read_excel(file, header=None, engine="openpyxl")
    BRIEF_ROW, OVERS_ROW = 1, 4

    meta: dict[str, dict[str, object]] = {}
    for col_idx in range(len(KEY_COLS), raw.shape[1]):
        item_code = str(raw.iloc[6, col_idx])  # codes live in row 7 (= index 6)
        if item_code == "nan":
            continue
        meta[item_code] = {
            "description": raw.iloc[BRIEF_ROW, col_idx],
            "overs": 0 if pd.isna(raw.iloc[OVERS_ROW, col_idx]) else raw.iloc[OVERS_ROW, col_idx],
        }
    return df, meta


def merge_allocations(dfs):
    """Outer‑merge *all* rows then aggregate by Store Number.

    * For key/string columns ⇒ first non‑null value.
    * For numeric item columns ⇒ sum (they're never duplicated across files, so
      sum equals the non‑null value; if you do have duplicates the numbers add).
    """
    combined = pd.concat(dfs, ignore_index=True, sort=False)

    # Coerce all non‑key columns to numeric where possible (item columns)
    non_key_cols = [c for c in combined.columns if c not in KEY_COLS]
    combined[non_key_cols] = combined[non_key_cols].apply(pd.to_numeric, errors="coerce")

    agg_funcs: dict[str, str] = {}
    for col in combined.columns:
        if col == "Store Number":
            continue
        agg_funcs[col] = "first" if col in KEY_COLS else "sum"

    master = combined.groupby("Store Number", as_index=False).agg(agg_funcs)
    master = master.sort_values("Store Number").reset_index(drop=True)
    return master


def write_with_metadata(master_df: pd.DataFrame, meta: dict[str, dict[str, object]]) -> BytesIO:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        STARTROW = 6  # leave rows 1‑6 free for metadata
        master_df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)
        ws = writer.sheets["Master Allocation"]

        item_cols = [c for c in master_df.columns if c not in KEY_COLS]
        for r_off, label in enumerate(LABELS):
            row_xl = 2 + r_off  # rows 2‑5
            ws.cell(row=row_xl, column=LABEL_COL_XL, value=label)
            for ic, item in enumerate(item_cols):
                col_xl = ITEM_START_XL + ic
                overs_val = meta.get(item, {}).get("overs", 0)
                total_alloc = master_df[item].fillna(0).sum()
                if r_off == 0:  # Brief Description
                    ws.cell(row=row_xl, column=col_xl, value=meta.get(item, {}).get("description", ""))
                elif r_off == 1:  # Total (inc Overs)
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc + overs_val)
                elif r_off == 2:  # Total Allocations
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc)
                elif r_off == 3:  # Overs
                    ws.cell(row=row_xl, column=col_xl, value=overs_val)
    buffer.seek(0)
    return buffer

# ──────────────────────────────────────────────────────────────────────────────
#  Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Allocation Merger", layout="wide")

st.title("Media Centre Allocation Merger")

st.markdown(
    """
    Upload one or more allocation exports → get a single consolidated workbook
    **with full store details**.
    """
)

uploaded_files = st.file_uploader(
    "Upload exports (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("👆 Drag your files here to begin.")
    st.stop()

progress = st.progress(0, text="Reading files…")
all_dfs, meta_dict = [], defaultdict(dict)
for i, up in enumerate(uploaded_files, start=1):
    df_part, meta_part = extract_data_and_meta(up)
    all_dfs.append(df_part)
    for k, v in meta_part.items():
        meta_dict.setdefault(k, v)  # keep first description/overs seen
    progress.progress(i / len(uploaded_files), text=f"Processed {i}/{len(uploaded_files)} file(s)")
progress.empty()

master_df = merge_allocations(all_dfs)

buffer = write_with_metadata(master_df, meta_dict)

st.success(
    f"✅ Merged {len(uploaded_files)} file{'s' if len(uploaded_files) > 1 else ''}: "
    f"{master_df.shape[0]} stores × {len(master_df.columns) - len(KEY_COLS)} items."
)

st.dataframe(master_df.head(50), use_container_width=True)

st.download_button(
    "📥 Download master_allocation.xlsx",
    data=buffer,
    file_name="master_allocation.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
