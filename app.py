import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Allocation Merger", layout="wide")

st.title("Media Centre Allocation Merger")

st.markdown(
    """
    Upload one or more **Media Centre** allocation exports (XLSX).
    The app will merge them into a single master allocation while:
    - Using the row with **Store Number** as header (row 7 in the raw export)
    - Aligning stores across files (outer join on Store Number)
    - Preserving all item allocation columns
    - Keeping the first non‑blank value for common store‑detail fields
    """
)

uploaded_files = st.file_uploader(
    "Upload allocation Excel files",
    type=["xlsx"],
    accept_multiple_files=True
)

@st.cache_data(show_spinner=False)
def load_file(file):
    """Load a single allocation export with the correct header row."""
    df = pd.read_excel(file, header=6)  # 0‑based row index; header row is row 7 in Excel
    # force Store Number to string to keep leading zeros, etc.
    df['Store Number'] = df['Store Number'].astype(str)
    return df

def merge_allocations(dfs):
    master = None
    # core store‑detail columns are always in positions 0‑10
    key_cols = [
        'Store Number', 'Store Name', 'Address Line 1', 'Address Line 2',
        'City or Town', 'County', 'Country', 'Post Code',
        'Region / Area', 'Location Type', 'Trading Format'
    ]
    for df in dfs:
        # split store detail vs item allocation columns
        item_cols = [c for c in df.columns if c not in key_cols]
        if master is None:
            master = df.copy()
        else:
            # perform outer join on Store Number only to avoid duplicate key columns
            merge_cols = ['Store Number']
            temp = df[merge_cols + item_cols]
            master = master.merge(temp, on='Store Number', how='outer')
    # Sort by Store Number for readability
    master = master.sort_values('Store Number').reset_index(drop=True)
    return master

if uploaded_files:
    dfs = [load_file(f) for f in uploaded_files]
    master_df = merge_allocations(dfs)

    st.success(f"Combined {len(uploaded_files)} file(s). "
               f"Master allocation has {master_df.shape[0]} stores "
               f"and {master_df.shape[1] - 11} item columns.")

    # Show a preview – the first 50 rows
    st.dataframe(master_df.head(50), use_container_width=True)

    # Provide download button
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        master_df.to_excel(writer, index=False, sheet_name="Master Allocation")
    buffer.seek(0)

    st.download_button(
        label="📥 Download Master Allocation",
        data=buffer,
        file_name="master_allocation.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("👆 Upload one or more Excel exports to begin.")
