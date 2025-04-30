import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

# Excel styling
try:
    import openpyxl
    from openpyxl.styles import Alignment, Border, Side, PatternFill, Font
    from openpyxl.utils import get_column_letter
except ImportError:
    st.error("❌ `openpyxl` is not installed. Please run `pip install openpyxl`.")
    st.stop()

# ───────────────────────── CONSTANTS ──────────────────────────
KEY_COLS = [
    "Store Number", "Store Name", "Address Line 1", "Address Line 2", "City or Town",
    "County", "Country", "Post Code", "Region / Area", "Location Type", "Trading Format",
]
LABELS = [
    "POS Code", "Kit Name", "Project Description", "Part", "Supplier",
    "Brief Description", "Total (inc Overs)", "Total Allocations", "Overs",
]
LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1   # column K (1-based)
ITEM_START_XL = LABEL_COL_XL + 1                       # first item col → L

# styles
THIN_SIDE = Side(style="thin", color="000000")
THIN_BORDER = Border(top=THIN_SIDE, left=THIN_SIDE, right=THIN_SIDE, bottom=THIN_SIDE)
ORANGE_FILL = PatternFill(start_color="F4B084", end_color="F4B084", fill_type="solid")
BOLD_FONT = Font(bold=True)

# ──────────────────────── DATA HELPERS ────────────────────────

def extract_alloc(file):
    """Return (df, meta) from one allocation export."""
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
    comb = pd.concat(dfs, ignore_index=True, sort=False)
    num_cols = [c for c in comb.columns if c not in KEY_COLS]
    comb[num_cols] = comb[num_cols].apply(pd.to_numeric, errors="coerce")
    agg = {c: ("first" if c in KEY_COLS else "sum") for c in comb.columns if c != "Store Number"}
    master = comb.groupby("Store Number", as_index=False).agg(agg)
    return master.sort_values("Store Number").reset_index(drop=True)


def load_brief(file):
    if file is None:
        return {}
    brief = pd.read_excel(file, header=1, engine="openpyxl")
    need = {"Brief Ref", "POS Code", "Project Description", "Part", "Supplier"}
    if miss := need - set(brief.columns):
        st.error("Consolidated Brief missing columns: " + ", ".join(miss))
        return {}
    out = {}
    for _, r in brief[list(need)].dropna(subset=["Brief Ref"]).iterrows():
        ref = str(r["Brief Ref"]).strip()
        out.setdefault(ref, {
            "pos_code": r["POS Code"],
            "project_description": r["Project Description"],
            "part": r["Part"],
            "supplier": r["Supplier"],
        })
    return out

# ──────────────────────── EXCEL WRITER ────────────────────────

def build_workbook(df: pd.DataFrame, meta: dict, event_code: str) -> BytesIO:
    """Return an in‑memory xlsx with full formatting."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        STARTROW = len(LABELS) + 1  # data header on Excel row 11
        df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)
        ws = writer.sheets["Master Allocation"]

        # Project Ref row
        ws.cell(row=1, column=1, value="Project Ref").font = BOLD_FONT
        ws.cell(row=1, column=2, value=event_code).font = BOLD_FONT

        # column widths & hide C-J
        for c in range(1, ws.max_column + 1):
            letter = get_column_letter(c)
            ws.column_dimensions[letter].width = 18
            if "C" <= letter <= "J":
                ws.column_dimensions[letter].hidden = True

        items = [c for c in df.columns if c not in KEY_COLS]

        # header rows 2‑10
        for r_off, label in enumerate(LABELS):
            row = 2 + r_off
            hdr_cell = ws.cell(row=row, column=LABEL_COL_XL, value=label)
            hdr_cell.alignment = Alignment(wrap_text=(row in (5, 7)), vertical="center")
            hdr_cell.fill = ORANGE_FILL
            hdr_cell.font = BOLD_FONT
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

        # style pandas header row (Excel row 11)
        excel_header_row = STARTROW + 1  # 0‑based to 1‑based
        for col in range(1, ws.max_column + 1):
            c = ws.cell(row=excel_header_row, column=col)
            c.fill = ORANGE_FILL
            c.font = BOLD_FONT
            c.border = THIN_BORDER

        # data rows borders & alignment
        data_start = excel_header_row + 1
        for row in ws.iter_rows(min_row=data_start, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                if cell.column >= ITEM_START_XL:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = THIN_BORDER

    buf.seek(0)
    return buf

# ────────────────────────── STREAMLIT UI ─────────────────────────
st.set_page_config(page_title="Superdrug Consolidated Allocation Builder", layout="wide")

st.title("Superdrug Consolidated Allocation Builder")

st.markdown(
    """
**Step 1 – Upload all allocation exports together** – [download them here](https://superdrug.aswmediacentre.com/ArtworkPrint/ArtworkPrintReport/ArtworkPrintReport?reportId=1149)  
**Step 2 – Upload the Consolidated Brief complete with Supplier for each line (optional)**  
**Step 3 – Enter the Event Code (required)**  
**Step 4 – Download the Consolidated Allocation**
    """
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

# combine allocations
progress = st.progress(0)
all_dfs, meta = [], defaultdict(dict)
for i, f in enumerate(alloc_files, 1):
    df_part, meta_part = extract_alloc(f)
    all_dfs.append(df_part)
    for k, v in meta_part.items():
        meta.setdefault(k, {}).update(v)
    progress.progress(i / len(alloc_files))
progress.empty()

# enrich from consolidated brief
for ref, info in load_brief(brief_file).items():
    meta.setdefault(ref, {}).update(info)

master_df = merge_allocations(all_dfs)

workbook_bytes = build_workbook(master_df, meta, event_code.strip())

st.success(
    f"Consolidated {master_df.shape[0]} stores × {len(master_df.columns) - len(KEY_COLS)} items."
