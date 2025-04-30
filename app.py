"""allocation_merger_app.py – v4
Adds the extra header rows (POS Code, Kit Name, Project Description, Part,
Supplier) above each item column and leaves their values blank for now.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

# ──────────────────────────────────────────────────────────────────────────────
#  Dependency check
# ──────────────────────────────────────────────────────────────────────────────
try:
    import openpyxl  # noqa: F401
except ImportError:
    st.error("❌ `openpyxl` is missing. Add it to *requirements.txt* or run `pip install openpyxl`.")
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

# Header rows we want above each item column (row numbers start at 2 in Excel)
LABELS = [
    "POS Code",            # row 2
    "Kit Name",            # row 3
    "Project Description", # row 4
    "Part",                # row 5
    "Supplier",            # row 6
    "Brief Description",   # row 7
    "Total (inc Overs)",   # row 8
    "Total Allocations",   # row 9
    "Overs",               # row 10
]

LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1  # K column (1‑based)
ITEM_START_XL = LABEL_COL_XL + 1                      # first item column L

# ──────────────────────────────────────────────────────────────────────────────
#  Helper functions
# ──────────────────────────────────────────────────────────────────────────────

def extract_data_and_meta(file):
    """Return (df, meta_dict) extracted from a single allocation export."""
    df = pd.read_excel(file, header=6, engine="openpyxl")
    df["Store Number"] = df["Store Number"].astype(str)

    raw = pd.read_excel(file, header=None, engine="openpyxl")
    BRIEF_ROW, OVERS_ROW = 1, 4  # relative to Excel (2nd & 5th rows visually)

    meta: dict[str, dict[str, object]] = {}
    for col_idx in range(len(KEY_COLS), raw.shape[1]):
        item_code = str(raw.iloc[6, col_idx])  # codes live in row 7 (index 6)
        if item_code == "nan":
            continue
        meta[item_code] = {
            "description": raw.iloc[BRIEF_ROW, col_idx],
            "overs": 0 if pd.isna(raw.iloc[OVERS_ROW, col_idx]) else raw.iloc[OVERS_ROW, col_idx],
        }
    return df, meta


def merge_allocations(dfs):
    """Merge all allocation data keeping first non-null store details."""
    combined = pd.concat(dfs, ignore_index=True, sort=False)
    non_key_cols = [c for c in combined.columns if c not in KEY_COLS]
    combined[non_key_cols] = combined[non_key_cols].apply(pd.to_numeric, errors="coerce")

    agg = {c: ("first" if c in KEY_COLS else "sum") for c in combined.columns if c != "Store Number"}
    master = combined.groupby("Store Number", as_index=False).agg(agg)
    master = master.sort_values("Store Number").reset_index(drop=True)
    return master


def write_with_metadata(master_df: pd.DataFrame, meta: dict[str, dict[str, object]]) -> BytesIO:
    """Create an in‑memory Excel workbook with all header rows populated."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        STARTROW = len(LABELS) + 2  # leave LABELS rows + one blank row
        master_df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)

        ws = writer.sheets["Master Allocation"]
        item_cols = [c for c in master_df.columns if c not in KEY_COLS]

        for r_off, label in enumerate(LABELS):
            row_xl = 2 + r_off  # rows 2 – 10
            ws.cell(row=row_xl, column=LABEL_COL_XL, value=label)
            for ic, item in enumerate(item_cols):
                col_xl = ITEM_START_XL + ic
                overs_val = meta.get(item, {}).get("overs", 0)
                total_alloc = master_df[item].fillna(0).sum()

                # Populate only the rows we have data for; leave new header rows blank
                if label == "Brief Description":
                    ws.cell(row=row_xl, column=col_xl, value=meta.get(item, {}).get("description", ""))
                elif label == "Total (inc Overs)":
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc + overs_val)
                elif label == "Total Allocations":
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc)
                elif label == "Overs":
                    ws.cell(row=row_xl, column=col_xl, value=overs_val)
                # POS Code, Kit Name, Project Description, Part, Supplier → left blank

    buffer.seek(0)
    return buffer

# ──────────────────────────────────────────────────────────────────────────────
#  Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Allocation Merger", layout="wide")

st.title("Media Centre Allocation Merger")

st.markdown(
    """
    Upload your allocation exports — the app builds a consolidated workbook and
    now includes extra header rows (*POS Code*, *Kit Name*, *Project
    Description*, *Part*, *Supplier*) ready for data.
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
        meta_dict.setdefault(k, v)
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
