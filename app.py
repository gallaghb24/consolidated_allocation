"""allocation_merger_app.py – v9
Formatting polish
─────────────────
• Borders now also applied to *A‑K* portion of the store table (rows 11‑end).  
• Header row 11 shaded light orange (#F4B084) and bold.  
• Label cells **K2‑K10** shaded the same orange & bold.  
• *A1* & *B1* bold.  
• Column width bumped to 18 chars (~125 px).  
• *Store Number* kept as an **integer** → sheet now sorts 3 → 9 → 10 → 1000.  
(no leading‑zero stores in source, so numeric safe)
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ `openpyxl` missing – add via pip."); st.stop()

# ─────────────────────────────────────────────
# Constants / styles
# ─────────────────────────────────────────────
KEY_COLS = [
    "Store Number","Store Name","Address Line 1","Address Line 2","City or Town",
    "County","Country","Post Code","Region / Area","Location Type","Trading Format",
]
LABELS = [
    "POS Code","Kit Name","Project Description","Part","Supplier",
    "Brief Description","Total (inc Overs)","Total Allocations","Overs",
]
LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1  # column K
ITEM_START_XL = LABEL_COL_XL + 1                      # column L
thin_side = Side(style="thin", color="000000")
THIN_BORDER = Border(top=thin_side, left=thin_side, right=thin_side, bottom=thin_side)
ORANGE_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
BOLD_FONT = Font(bold=True)

# ─────────────────────────────────────────────
# Data helpers
# ─────────────────────────────────────────────

def extract_alloc_data_and_meta(file):
    df = pd.read_excel(file, header=6, engine="openpyxl")
    # Ensure Store Number numeric (drops .0 floats)
    df["Store Number"] = pd.to_numeric(df["Store Number"], errors="coerce").astype("Int64")
    raw = pd.read_excel(file, header=None, engine="openpyxl")
    meta = {}
    for col in range(len(KEY_COLS), raw.shape[1]):
        ref = str(raw.iloc[6, col])
        if ref == "nan":
            continue
        meta[ref] = {
            "brief_description": raw.iloc[1, col],
            "overs": 0 if pd.isna(raw.iloc[4, col]) else raw.iloc[4, col],
        }
    return df, meta

def merge_allocations(dfs):
    combined = pd.concat(dfs, ignore_index=True, sort=False)
    numeric_cols = [c for c in combined.columns if c not in KEY_COLS]
    combined[numeric_cols] = combined[numeric_cols].apply(pd.to_numeric, errors="coerce")
    agg = {c: ("first" if c in KEY_COLS else "sum") for c in combined.columns if c != "Store Number"}
    master = combined.groupby("Store Number", as_index=False).agg(agg).sort_values("Store Number").reset_index(drop=True)
    return master

def load_consolidated_brief(file):
    if file is None:
        return {}
    cb = pd.read_excel(file, header=1, engine="openpyxl")
    needed = {"Brief Ref","POS Code","Project Description","Part","Supplier"}
    if missing := needed - set(cb.columns):
        st.error("Consolidated Brief missing: "+", ".join(missing)); return {}
    out = {}
    for _, r in cb[list(needed)].dropna(subset=["Brief Ref"]).iterrows():
        ref = str(r["Brief Ref"]).strip()
        out.setdefault(ref, {
            "pos_code": r["POS Code"],
            "project_description": r["Project Description"],
            "part": r["Part"],
            "supplier": r["Supplier"],
        })
    return out

# ─────────────────────────────────────────────
# Excel writer
# ─────────────────────────────────────────────

def write_with_metadata(df, meta, event):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        STARTROW = len(LABELS) + 1  # Excel row 11 is header
        df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)
        ws = writer.sheets["Master Allocation"]

        # Project Ref row
        ws.cell(row=1, column=1, value="Project Ref").font = BOLD_FONT
        ws.cell(row=1, column=2, value=event).font = BOLD_FONT

        # Column widths & hide C‑J
        for col in range(1, ws.max_column + 1):
            let = get_column_letter(col)
            ws.column_dimensions[let].width = 18  # wider
            if "C" <= let <= "J":
                ws.column_dimensions[let].hidden = True

        items = [c for c in df.columns if c not in KEY_COLS]
        for r_off, label in enumerate(LABELS):
            row = 2 + r_off
            label_cell = ws.cell(row=row, column=LABEL_COL_XL, value=label)
            label_cell.alignment = Alignment(wrap_text=(row in (5, 7)), vertical="center")
            if LABEL_COL_XL == 11:  # K column
                label_cell.fill = ORANGE_FILL
                label_cell.font = BOLD_FONT
            for i, item in enumerate(items):
                col_xl = ITEM_START_XL + i
                cell = ws.cell(row=row, column=col_xl)
                data = meta.get(item, {})
                overs = data.get("overs", 0)
                total = df[item].fillna(0).sum()
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
                    cell.value = total + overs
                elif label == "Total Allocations":
                    cell.value = total
                elif label == "Overs":
                    cell.value = overs
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(row in (5, 7)))
                cell.border = THIN_BORDER
        # Style pandas header row (Excel row 11)
        header_row = STARTROW  # 0‑based → Excel row 11
        for col in range(1, ws.max_column + 1):
            c = ws.cell(row=header_row + 1, column=col)
            c.fill = ORANGE_FILL
            c.font = BOLD_FONT
            c.border = THIN_BORDER
        # Data rows
        data_start = header_row + 2  # first data row
        for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.column >= ITEM_START_XL:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = THIN_BORDER
        # Extra border for A‑K (columns 1‑11) already covered by loop above
    buf.seek(0); return buf

# ─────────────────────────────────────────────
# Streamlit UI (title & steps unchanged from v8)
# ─────────────────────────────────────────────
st.set_page_config(page_title="Superdrug Consolidated Allocation Builder", layout="wide")

st.title("Superdrug Consolidated Allocation Builder")

st.markdown(
    """
    **Step 1 – Upload all allocation exports together** – [download them here](https://superdrug.aswmediacentre.com/ArtworkPrint/ArtworkPrintReport/ArtworkPrintReport?reportId=1149)  
    **Step 2 – Upload the Consolidated Brief** (optional)  
    **Step 3 – Enter the Event Code** (required)  
    **Step 4 – Download the Consolidated Allocation**
    """
)

alloc_files = st.file_uploader("Allocation exports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
brief_file = st.file_uploader("Consolidated Brief (.xlsx)", type=["xlsx"], key="brief")
event_code = st.text_input("Event Code (e.g. E0625)")

if not alloc_files:
    st.info("Please upload at least one allocation export."); st.stop()
if not event_code.strip():
    st.warning("Event Code is required."); st.stop()

# Process uploads
prog = st.progress(0)
all_dfs, meta = [], defaultdict(dict)
for i, f in enumerate(alloc_files, start=1):
    d, m = extract_alloc_data_and_meta(f)
    all_dfs.append(d)
    for k, v in m.items():
        meta.setdefault(k, {}).update(v)
    prog.progress(i / len(alloc_files))
prog.empty()

for ref, info in load_consolidated_brief(brief_file).items():
    meta.setdefault(ref, {}).update(info)

master = merge_allocations(all_dfs)
file_bytes = write_with_metadata(master, meta, event_code.strip())

st.success(f"Consolidated {master.shape[0]} stores × {len(master.columns) - len(KEY_COLS)} items.")

st.dataframe(master.head(50), use_container_width=True)

st.download_button("📥 Download the Consolidated Allocation", data=file_bytes, file_name=f"{event_code.strip()}_Consolidated_Allocation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
