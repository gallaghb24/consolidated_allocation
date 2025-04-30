"""allocation_merger_app.py – v2
A Streamlit tool that merges Media Centre allocation exports **and** keeps the
metadata rows (Brief Description, Total (inc Overs), Total Allocations, Overs)
above each item column.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

# ──────────────────────────────────────────────────────────────────────────────
#  Dependencies check
# ──────────────────────────────────────────────────────────────────────────────
try:
    import openpyxl  # noqa: F401 – so pandas can use the engine
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error(
        "❌ `openpyxl` is missing. Add it to *requirements.txt* or run "
        "`pip install openpyxl` and restart the app."
    )
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
#  Constants
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

# Column in which the row labels (above item columns) live – same as "Trading Format"
LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1  # 1‑based for openpyxl
ITEM_START_XL = LABEL_COL_XL + 1  # first item column (e.g. L)

# ──────────────────────────────────────────────────────────────────────────────
#  Helper functions
# ──────────────────────────────────────────────────────────────────────────────

def extract_data_and_meta(file):
    """Return (df, meta_dict) for a single allocation export."""

    # 1️⃣  Store‑level allocations (row 7 is the header)
    df = pd.read_excel(file, header=6, engine="openpyxl")
    df["Store Number"] = df["Store Number"].astype(str)

    # 2️⃣  Metadata rows (read raw)
    raw = pd.read_excel(file, header=None, engine="openpyxl")

    # Row indices (0‑based) relative to Excel file layout
    BRIEF_ROW = 1
    OVERS_ROW = 4

    meta: dict[str, dict[str, object]] = {}
    for col_idx in range(len(KEY_COLS), raw.shape[1]):
        item_code = str(raw.iloc[6, col_idx])  # row 7 (index 6) contains codes
        if item_code == "nan":  # skip empty tail columns
            continue
        description = raw.iloc[BRIEF_ROW, col_idx]
        overs_val = raw.iloc[OVERS_ROW, col_idx]
        overs_val = 0 if pd.isna(overs_val) else overs_val

        meta[item_code] = {
            "description": description,
            "overs": overs_val,
        }

    return df, meta


def merge_allocations(dfs):
    """Outer‑join all dfs on *Store Number* and keep every item column."""
    master: pd.DataFrame | None = None
    for df in dfs:
        item_cols = [c for c in df.columns if c not in KEY_COLS]
        if master is None:
            master = df.copy()
        else:
            temp = df[["Store Number"] + item_cols]
            master = master.merge(temp, on="Store Number", how="outer")

    master = master.sort_values("Store Number").reset_index(drop=True)
    return master


def write_with_metadata(master_df: pd.DataFrame, meta: dict[str, dict[str, object]]) -> BytesIO:
    """Return an in‑memory Excel file containing metadata rows + data."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # Write store allocations starting on row 7 (Excel) so we have 6 rows for metadata
        STARTROW = 6  # 0‑based for pandas → row 7 in Excel
        master_df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)

        ws = writer.sheets["Master Allocation"]
        item_cols = [c for c in master_df.columns if c not in KEY_COLS]

        for row_offset, label in enumerate(LABELS):
            row_xl = 2 + row_offset  # Excel rows 2‑5
            ws.cell(row=row_xl, column=LABEL_COL_XL, value=label)

            for ic, item in enumerate(item_cols):
                col_xl = ITEM_START_XL + ic  # 1‑based Excel col number
                if row_offset == 0:  # Brief Description
                    ws.cell(row=row_xl, column=col_xl, value=meta.get(item, {}).get("description", ""))
                elif row_offset == 3:  # Overs
                    ws.cell(row=row_xl, column=col_xl, value=meta.get(item, {}).get("overs", 0))
                elif row_offset == 2:  # Total Allocations
                    total_alloc = master_df[item].fillna(0).sum()
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc)
                elif row_offset == 1:  # Total (inc Overs)
                    overs_val = meta.get(item, {}).get("overs", 0)
                    total_alloc = master_df[item].fillna(0).sum()
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc + overs_val)

    buffer.seek(0)
    return buffer

# ──────────────────────────────────────────────────────────────────────────────
#  Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Allocation Merger", layout="wide")

st.title("Media Centre Allocation Merger")

st.markdown(
    """
    **How it works**  
    1. Drop in one or more Media Centre allocation exports (`.xlsx`).  
    2. The app outer‑joins on **Store Number** so any new stores are kept.  
    3. All item columns are preserved.  
    4. The rows above each item column – *Brief Description*, *Overs*, etc. – are
       reproduced in the final file.  
    5. *Total Allocations* is re‑calculated to guarantee it matches the table,
       and *Total (inc Overs)* = Total Allocations + Overs.
    """
)

uploaded_files = st.file_uploader(
    "Upload allocation Excel files",
    type=["xlsx"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("👆 Upload the exports to begin.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
#  Load & merge
# ──────────────────────────────────────────────────────────────────────────────
progress = st.progress(0, text="Reading files…")
all_dfs, all_meta_dicts = [], defaultdict(dict)
for i, file in enumerate(uploaded_files, start=1):
    df_part, meta_part = extract_data_and_meta(file)
    all_dfs.append(df_part)
    for k, v in meta_part.items():
        # in case of duplicates across files, first one wins → can be changed to validation/merge if needed
        all_meta_dicts.setdefault(k, v)
    progress.progress(i / len(uploaded_files), text=f"Processed {i}/{len(uploaded_files)} file(s)…")
progress.empty()

master_df = merge_allocations(all_dfs)

# ──────────────────────────────────────────────────────────────────────────────
#  Build Excel & offer download
# ──────────────────────────────────────────────────────────────────────────────
buffer = write_with_metadata(master_df, all_meta_dicts)

st.success(
    f"Combined **{len(uploaded_files)}** file{'s' if len(uploaded_files) > 1 else ''}.  "
    f"Master allocation: **{master_df.shape[0]}** stores × **{len(master_df.columns) - len(KEY_COLS)}** items."
)

st.dataframe(master_df.head(50), use_container_width=True)

st.download_button(
    "📥 Download master_allocation.xlsx",
    data=buffer,
    file_name="master_allocation.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
