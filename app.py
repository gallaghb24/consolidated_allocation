"""allocation_merger_app.py – v5
Adds:
* Second uploader for **Consolidated Brief** (expects exactly one file).
* Text input for mandatory **Event Code** (e.g. E0625).
* Populates POS Code, Project Description, Part, Supplier rows by matching each
  item column’s *Brief Ref* (e.g. SDG1849/39609) against the Consolidated
  Brief.
* Writes the Event Code into cell **B1** (labelled *Project Ref*) and uses it
  in the download filename (`<EventCode>_master_allocation.xlsx`).
* Kit Name row remains blank.
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from collections import defaultdict

# ──────────────────────────────────────────────────────────────────────────────
#  Dependency check
# ──────────────────────────────────────────────────────────────────────────────
try:
    import openpyxl  # noqa: F401 – needed as the Excel engine
except ImportError:
    st.error("❌ `openpyxl` is missing. Add it to requirements.txt or `pip install openpyxl`.")
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
    "POS Code",            # row 2
    "Kit Name",            # row 3 (blank)
    "Project Description", # row 4
    "Part",                # row 5
    "Supplier",            # row 6
    "Brief Description",   # row 7
    "Total (inc Overs)",   # row 8
    "Total Allocations",   # row 9
    "Overs",               # row 10
]

LABEL_COL_XL = KEY_COLS.index("Trading Format") + 1  # K → 1‑based index
ITEM_START_XL = LABEL_COL_XL + 1                      # L, M, …

# ──────────────────────────────────────────────────────────────────────────────
#  Helper functions
# ──────────────────────────────────────────────────────────────────────────────

def extract_alloc_data_and_meta(file):
    """Return (df, meta) from a single allocation export."""
    df = pd.read_excel(file, header=6, engine="openpyxl")
    df["Store Number"] = df["Store Number"].astype(str)

    raw = pd.read_excel(file, header=None, engine="openpyxl")
    BRIEF_ROW, OVERS_ROW = 1, 4

    meta: dict[str, dict[str, object]] = {}
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
    """Merge keeping first non-null store-detail fields."""
    combined = pd.concat(dfs, ignore_index=True, sort=False)
    numeric_cols = [c for c in combined.columns if c not in KEY_COLS]
    combined[numeric_cols] = combined[numeric_cols].apply(pd.to_numeric, errors="coerce")

    agg = {c: ("first" if c in KEY_COLS else "sum") for c in combined.columns if c != "Store Number"}
    master = combined.groupby("Store Number", as_index=False).agg(agg)
    return master.sort_values("Store Number").reset_index(drop=True)


def load_consolidated_brief(file):
    """Read the Consolidated Brief → dict keyed by Brief Ref."""
    if file is None:
        return {}
    try:
        cb = pd.read_excel(file, header=1, engine="openpyxl")  # header row is row 2 in Excel
    except Exception as exc:
        st.error(f"Failed to read Consolidated Brief: {exc}")
        return {}

    cols_needed = {
        "Brief Ref": "brief_ref",
        "POS Code": "pos_code",
        "Project Description": "project_description",
        "Part": "part",
        "Supplier": "supplier",
    }
    missing = [c for c in cols_needed if c not in cb.columns]
    if missing:
        st.error(f"Consolidated Brief is missing expected columns: {', '.join(missing)}")
        return {}

    brief_dict = {}
    for _, row in cb[cols_needed.keys()].dropna(subset=["Brief Ref"]).iterrows():
        brief_ref = str(row["Brief Ref"]).strip()
        if brief_ref not in brief_dict:  # keep first occurrence
            brief_dict[brief_ref] = {
                "pos_code": row["POS Code"],
                "project_description": row["Project Description"],
                "part": row["Part"],
                "supplier": row["Supplier"],
            }
    return brief_dict


def write_with_metadata(master_df: pd.DataFrame, meta: dict[str, dict[str, object]], event_code: str) -> BytesIO:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        STARTROW = len(LABELS) + 2  # leave header rows and one blank spacer
        master_df.to_excel(writer, index=False, sheet_name="Master Allocation", startrow=STARTROW)
        ws = writer.sheets["Master Allocation"]

        # Row 1 – Project Ref + Event code
        ws.cell(row=1, column=1, value="Project Ref")
        ws.cell(row=1, column=2, value=event_code)

        item_cols = [c for c in master_df.columns if c not in KEY_COLS]

        for r_off, label in enumerate(LABELS):
            row_xl = 2 + r_off  # rows 2‑10
            ws.cell(row=row_xl, column=LABEL_COL_XL, value=label)
            for ic, item in enumerate(item_cols):
                col_xl = ITEM_START_XL + ic

                data = meta.get(item, {})
                overs_val = data.get("overs", 0)
                total_alloc = master_df[item].fillna(0).sum()

                if label == "POS Code":
                    ws.cell(row=row_xl, column=col_xl, value=data.get("pos_code", ""))
                elif label == "Kit Name":
                    pass  # remains blank
                elif label == "Project Description":
                    ws.cell(row=row_xl, column=col_xl, value=data.get("project_description", ""))
                elif label == "Part":
                    ws.cell(row=row_xl, column=col_xl, value=data.get("part", ""))
                elif label == "Supplier":
                    ws.cell(row=row_xl, column=col_xl, value=data.get("supplier", ""))
                elif label == "Brief Description":
                    ws.cell(row=row_xl, column=col_xl, value=data.get("brief_description", ""))
                elif label == "Total (inc Overs)":
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc + overs_val)
                elif label == "Total Allocations":
                    ws.cell(row=row_xl, column=col_xl, value=total_alloc)
                elif label == "Overs":
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
    1. **Upload** one or more *allocation exports* (XLSX).  
    2. **Upload** the single *Consolidated Brief* (XLSX) – used to fill the
       POS Code / Project Description / Part / Supplier rows.  
    3. **Enter** the Event Code (e.g. `E0625`).  
    4. Click *Download* to get your consolidated workbook and allocations.
    """
)

alloc_files = st.file_uploader(
    "Allocation export files (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=True,
)

brief_file = st.file_uploader(
    "Consolidated Brief (.xlsx)",
    type=["xlsx"],
    accept_multiple_files=False,
    key="brief_uploader",
)

event_code = st.text_input("Event Code (required)")

if not alloc_files:
    st.info("👆 Please upload at least one allocation export to begin.")
    st.stop()

if brief_file is None:
    st.warning("⬆️ Upload the Consolidated Brief to populate header rows (you can still merge without it, rows will stay blank).")

if not event_code.strip():
    st.warning("✏️ Enter the Event Code – needed for row 1 and filename.")
    st.stop()

a_progress = st.progress(0, text="Reading allocation files…")
all_dfs, meta = [], defaultdict(dict)
for i, up in enumerate(alloc_files, start=1):
    df_part, meta_part = extract_alloc_data_and_meta(up)
    all_dfs.append(df_part)
    for k, v in meta_part.items():
        meta.setdefault(k, {}).update(v)  # keep/merge description & overs
    a_progress.progress(i / len(alloc_files), text=f"Processed {i}/{len(alloc_files)} file(s)")
a_progress.empty()

# Load consolidated brief → enrich meta
brief_dict = load_consolidated_brief(brief_file)
for ref, info in brief_dict.items():
    meta.setdefault(ref, {}).update(info)

master_df = merge_allocations(all_dfs)

buffer = write_with_metadata(master_df, meta, event_code.strip())

download_name = f"{event_code.strip()}_master_allocation.xlsx"

st.success(
    f"✅ Created consolidated allocation: {master_df.shape[0]} stores × "
    f"{len(master_df.columns) - len(KEY_COLS)} items."
)

st.dataframe(master_df.head(50), use_container_width=True)

st.download_button(
    "📥 Download master allocation",
    data=buffer,
    file_name=download_name,
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
