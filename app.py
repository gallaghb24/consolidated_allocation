"""allocation_merger_app.py – v6
Adds final formatting tweaks requested:
1. Download filename → <EventCode>_Consolidated_Allocation.xlsx
2. Wrap text in rows 5 & 7 (Part, Brief Description)
3. Make *every* row ~100 px tall (≈ 75 pt)
4. Hide columns C‒J
5. Centre-align everything from column L onward
6. Apply a thin border around the header block (rows 2 – 10, K → last item col)
7. Apply the same thin border to the entire store allocation table below
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ `openpyxl` is missing. Add it to requirements.txt or run `pip install openpyxl`.")
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
    "POS Code",            # row 2
    "Kit Name",            # row 3 (blank)
    "Project Description", # row 4
    "Part",                # row 5
    "Supplier",            # row 6
    "Brief Description",   # row 7
    "Total (inc Overs)",   # row 8
    "Total Allocations",   # row 9
    "Overs",               # row 10
]

LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1  # → column K (1-based)
ITEM_START_XL = LABEL_COL_XL + 1                      # → first item col L
ROW_HEIGHT_PT = 75  # ≈ 100 px

thin_side = Side(style="thin", color="000000")
THIN_BORDER = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)

# ──────────────────────────────────────────────────────────────────────────────
#  Helper functions
# ──────────────────────────────────────────────────────────────────────────────

def extract_alloc_data_and_meta(file):
    df = pd.read_excel(file, header=6, engine="openpyxl")
    df["Store Number"] = df["Store Number"].astype(str)

    raw = pd.read_excel(file, header=None, engine="openpyxl")
    BRIEF_ROW, OVERS_ROW = 1, 4

    meta = {}
    for col_idx in range(len(KEY_COLS), raw.shape[1]):
        brief_ref = str(raw.iloc[6, col_idx])
        if brief_ref == "nan":
            continue
        meta[brief_ref] = {
            "brief_description": raw.iloc[BRIEF_ROW, col_idx],
            "overs": 0 if pd.isna(raw.iloc[OVERS_ROW, col_idx]) else raw.iloc[OVERS_ROW, col_idx],
        }
    return df, meta


def merge_allocations(dfs):
    combined = pd.concat(dfs, ignore_index=True, sort=False)
    numeric_cols = [c for c in combined.columns if c not in KEY_COLS]
    combined[numeric_cols] = combined[numeric_cols].apply(pd.to_numeric, errors="coerce")
    agg = {c: ("first" if c in KEY_COLS else "sum") for c in combined.columns if c != "Store Number"}
    master = combined.groupby("Store Number", as_index=False).agg(agg)
    return master.sort_values("Store Number").reset_index(drop=True)


def load_consolidated_brief(file):
    if file is None:
        return {}
    cb = pd.read_excel(file, header=1, engine="openpyxl")
    cols = {
        "Brief Ref": "brief_ref",
        "POS Code": "pos_code",
        "Project Description": "project_description",
        "Part": "part",
        "Supplier": "supplier",
    }
    if missing := [c for c in cols if c not in cb.columns]:
        st.error("Consolidated Brief missing columns: " + ", ".join(missing))
        return {}
    brief_dict = {}
    for _, row in cb[cols.keys()].dropna(subset=["Brief Ref"]).iterrows():
        ref = str(row["Brief Ref"]).strip()
        brief_dict.setdefault(ref, {
            "pos_code": row["POS Code"],
            "project_description": row["Project Description"],
            "part": row["Part"],
            "supplier": row["Supplier"],
        })
    return brief_dict


def write_with_metadata(master_df, meta, event_code):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        STARTROW = len(LABELS) + 2  # leaves one blank row after headers
        master_df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)
        ws = writer.sheets["Master Allocation"]

        # Row 1: Project Ref + Event code
        ws.cell(row=1, column=1, value="Project Ref")
        ws.cell(row=1, column=2, value=event_code)

        # Column setup: set width & hide C-J
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].width = 14.3  # ≈100 px
            if "C" <= letter <= "J":
                ws.column_dimensions[letter].hidden = True

        # Header labels + data
        item_cols = [c for c in master_df.columns if c not in KEY_COLS]
        for r_off, label in enumerate(LABELS):
            row_xl = 2 + r_off  # rows 2-10
            ws.cell(row=row_xl, column=LABEL_COL_XL, value=label)
            for ic, item in enumerate(item_cols):
                col_xl = ITEM_START_XL + ic
                cell = ws.cell(row=row_xl, column=col_xl)

                # Fill data where we have it
                data = meta.get(item, {})
                overs_val = data.get("overs", 0)
                total_alloc = master_df[item].fillna(0).sum()
                if label == "POS Code":
                    cell.value = data.get("pos_code", "")
                elif label == "Project Description":
                    cell.value = data.get("project_description", "")
                elif label == "Part":
                    cell.value = data.get("part", "")
                elif label == "Supplier":
                    cell.value = data.get("supplier", "")
                elif label == "Brief Description":
                    cell.value = data.get("brief_description", "")
                elif label == "Total (inc Overs)":
                    cell.value = total_alloc + overs_val
                elif label == "Total Allocations":
                    cell.value = total_alloc
                elif label == "Overs":
                    cell.value = overs_val

                # Centring for item columns
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(row_xl in (5, 7)))

            # Label column alignment & wrapping
            label_cell = ws.cell(row=row_xl, column=LABEL_COL_XL)
            label_cell.alignment = Alignment(wrap_text=(row_xl in (5, 7)), vertical="center")

        # Row height for all rows
        for row in range(1, ws.max_row + 1):
            ws.row_dimensions[row].height = ROW_HEIGHT_PT

        # Centre-align item columns in the data table & apply borders
        table_min_row = STARTROW + 1  # header row of DF
        for row in ws.iter_rows(min_row=table_min_row, max_row=ws.max_row, min_col=ITEM_START_XL, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # Borders for header block
        for row in ws.iter_rows(min_row=2, max_row=10, min_col=LABEL_COL_XL, max_col=ws.max_column):
            for cell in row:
                cell.border = THIN_BORDER

        # Borders for store allocation table
        for row in ws.iter_rows(min_row=table_min_row, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.border = THIN_BORDER

    buffer.seek(0)
    return buffer

# ──────────────────────────────────────────────────────────────────────────────
#  Streamlit UI
# ──────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Allocation Merger", layout="wide")

st.title("Media Centre Allocation Merger")

st.markdown(
    """
    1. Upload one or more **allocation exports**.  
    2. Upload the **Consolidated Brief** (1 file).  
    3. Enter the **Event Code** (required).  
    4. Download the formatted workbook.
    """
)

alloc_files = st.file_uploader("Allocation exports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
brief_file = st.file_uploader("Consolidated Brief (.xlsx)", type=["xlsx"], accept_multiple_files=False, key="brief_uploader")
event_code = st.text_input("Event Code (e.g. E0625)")

if not alloc_files:
    st.info("👆 Please upload at least one allocation export to begin.")
    st.stop()
if not event_code.strip():
    st.warning("✏️ Enter the Event Code – required for export.")
    st.stop()

# ──────────────────────────────────────────────────────────────────────────────
#  Processing
# ──────────────────────────────────────────────────────────────────────────────
a_progress = st.progress(0, text="Reading allocation files…")
all_dfs, meta = [], defaultdict(dict)
for i, up in enumerate(alloc_files, start=1):
    df_part, meta_part = extract_alloc_data_and_meta(up)
    all_dfs.append(df_part)
    for k, v in meta_part.items():
        meta.setdefault(k, {}).update(v)
    a_progress.progress(i / len(alloc_files), text=f"Processed {i}/{len(alloc_files)} file(s)")
a_progress.empty()

brief_dict = load_consolidated_brief(brief_file)
for ref, info in brief_dict.items():
    meta.setdefault(ref, {}).update(info)

master_df = merge_allocations(all_dfs)

buffer = write_with_metadata(master_df, meta, event_code.strip())

download_name = f"{event_code.strip()}_Consolidated_Allocation.xlsx"

st.success(
    f"✅ Consolidated: {master_df.shape[0]} stores × {len(master_df.columns) - len(KEY_COLS)} items."
)

st.dataframe(master_df.head(50), use_container_width=True)

st.download_button("📥 Download consolidated allocation", data=buffer, file_name=download_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
