import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

# ───────────────────────────── Excel styling helpers ─────────────────────────────
try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ `openpyxl` is not installed. Please run `pip install openpyxl`.")
    st.stop()

THIN_SIDE = Side(style="thin", color="000000")
THIN_BORDER = Border(top=THIN_SIDE, left=THIN_SIDE, right=THIN_SIDE, bottom=THIN_SIDE)
ORANGE_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
BOLD_FONT = Font(bold=True)

# ───────────────────────────── Column definitions ─────────────────────────────
KEY_COLS = [
    "Store Number", "Store Name", "Address Line 1", "Address Line 2", "City or Town",
    "County", "Country", "Post Code", "Region / Area", "Location Type", "Trading Format",
]
LABELS = [
    "POS Code", "Kit Name", "Project Description", "Part", "Supplier",
    "Brief Description", "Total (inc Overs)", "Total Allocations", "Overs",
]
LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1  # column K (1‑based)
ITEM_START_XL = LABEL_COL_XL + 1                     # column L

# ───────────────────────────── Data helpers ─────────────────────────────

def extract_alloc(file):
    """Read a single allocation export and return (df, meta)."""
    df = pd.read_excel(file, header=6, engine="openpyxl")
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
    if not dfs:
        return pd.DataFrame()
    combined = pd.concat(dfs, ignore_index=True, sort=False)
    num_cols = [c for c in combined.columns if c not in KEY_COLS]
    combined[num_cols] = combined[num_cols].apply(pd.to_numeric, errors="coerce")
    agg = {c: ("first" if c in KEY_COLS else "sum") for c in combined.columns if c != "Store Number"}
    return combined.groupby("Store Number", as_index=False).agg(agg).sort_values("Store Number").reset_index(drop=True)


def load_brief(file):
    """Read consolidated brief → dict keyed by Brief Ref."""
    if file is None:
        return {}
    brief = pd.read_excel(file, header=1, engine="openpyxl")
    required = {"Brief Ref", "POS Code", "Project Description", "Part", "Supplier"}
    missing = required - set(brief.columns)
    if missing:
        st.error("Consolidated Brief missing columns: " + ", ".join(missing))
        return {}
    out = {}
    for _, row in brief[list(required)].dropna(subset=["Brief Ref"]).iterrows():
        ref = str(row["Brief Ref"]).strip()
        out.setdefault(ref, {
            "pos_code": row["POS Code"],
            "project_description": row["Project Description"],
            "part": row["Part"],
            "supplier": row["Supplier"],
        })
    return out

# ───────────────────────────── Workbook builder ─────────────────────────────

def build_workbook(df: pd.DataFrame, meta: dict, event_code: str) -> BytesIO:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        STARTROW = len(LABELS) + 1  # data header goes on Excel row 11
        df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)
        ws = writer.sheets["Master Allocation"]

        # Row 1 – Project Ref & Event Code
        ws.cell(row=1, column=1, value="Project Ref").font = BOLD_FONT
        ws.cell(row=1, column=2, value=event_code).font = BOLD_FONT

        # Column widths & hide C–J
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            ws.column_dimensions[letter].width = 18
            if "C" <= letter <= "J":
                ws.column_dimensions[letter].hidden = True

        item_cols = [c for c in df.columns if c not in KEY_COLS]

        # Rows 2–10 labels + values
        for r_off, label in enumerate(LABELS):
            row_num = 2 + r_off
            label_cell = ws.cell(row=row_num, column=LABEL_COL_XL, value=label)
            label_cell.alignment = Alignment(wrap_text=(row_num in (5, 7)), vertical="center")
            label_cell.fill = ORANGE_FILL
            label_cell.font = BOLD_FONT
            for i, item in enumerate(item_cols):
                cell = ws.cell(row=row_num, column=ITEM_START_XL + i)
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
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=(row_num in (5, 7)))
                cell.border = THIN_BORDER

        # Style pandas-generated header (Excel row 11)
        header_excel_row = STARTROW + 1
        for col_idx in range(1, ws.max_column + 1):
            c = ws.cell(row=header_excel_row, column=col_idx)
            c.fill = ORANGE_FILL
            c.font = BOLD_FONT
            c.border = THIN_BORDER

        # Data rows borders & alignment
        data_start_row = header_excel_row + 1
        for row in ws.iter_rows(min_row=data_start_row, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.column >= ITEM_START_XL:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = THIN_BORDER

    buf.seek(0)
    return buf

# ───────────────────────────── Streamlit UI ─────────────────────────────

st.set_page_config(page_title="Superdrug Consolidated Allocation Builder", layout="wide")

st.title("Superdrug Consolidated Allocation Builder")

st.markdown(
    """**Step 1 – Upload all allocation exports together** – [download them here](https://superdrug.aswmediacentre.com/ArtworkPrint/ArtworkPrintReport/ArtworkPrintReport?reportId=1149)  
**Step 2 – Upload the Consolidated Brief complete with Supplier for each line (optional)**  
**Step 3 – Enter the Event Code (required)**  
**Step 4 – Download the Consolidated Allocation**"""
)

alloc_files = st.file_uploader("Allocation exports (.xlsx)", type=["xlsx"], accept_multiple_files=True)
brief_file = st.file_uploader("Consolidated Brief (.xlsx)", type=["xlsx"], key="brief")
event_code = st.text_input("Event Code (e.g. E0625)")

if not alloc_files:
    st.info("Please upload at least one allocation export.")
    st.stop()
if not event_code.strip():
    st.warning("Event Code is required.")
    st.stop()

# ───────────────────────────── Merge process ─────────────────────────────
progress = st.progress(0)
all_dfs, meta = [], defaultdict(dict)
for idx, upload in enumerate(alloc_files, start=1):
    part_df, part_meta = extract_alloc(upload)
    all_dfs.append(part_df)
    for key, val in part_meta.items():
        meta.setdefault(key, {}).update(val)
    progress.progress(idx / len(alloc_files))
progress.empty()

# Enrich metadata from brief
for ref, info in load_brief(brief_file).items():
    meta.setdefault(ref, {}).update(info)

master_df = merge_allocations(all_dfs)

workbook = build_workbook(master
