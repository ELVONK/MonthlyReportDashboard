# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import re, time, uuid
import xlsxwriter

st.set_page_config(page_title="Monthly Report Dashboard", layout="wide")
st.title("Monthly Report Dashboard — Upload, Preview & Auto-Named Reports")

# =====================
# Paths & Utilities
# =====================
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
for d in (UPLOAD_DIR, REPORT_DIR):
    d.mkdir(parents=True, exist_ok=True)


def safe_unique_name(original_name: str) -> str:
    p = Path(original_name)
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", p.stem)[:80] or "file"
    suffix = p.suffix.lower()
    ts = time.strftime("%Y%m%d-%H%M%S")
    uid = uuid.uuid4().hex[:6]
    return f"{stem}_{ts}_{uid}{suffix}"


def save_uploaded_file(up_file) -> Path:
    dest = UPLOAD_DIR / safe_unique_name(up_file.name)
    with open(dest, "wb") as f:
        f.write(up_file.getbuffer())
    return dest


def list_saved_files(dir_path: Path) -> pd.DataFrame:
    rows = []
    for p in sorted(dir_path.glob("*")):
        if p.is_file():
            rows.append({
                "file": p.name,
                "size_kb": round(p.stat().st_size / 1024, 1),
                "modified": time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(p.stat().st_mtime)),
            })
    return pd.DataFrame(rows)


# =====================
# Caching helpers
# =====================
@st.cache_data(show_spinner=False)
def read_csv_preview(file_bytes: bytes, nrows: int = 15) -> pd.DataFrame:
    from io import BytesIO
    return pd.read_csv(BytesIO(file_bytes), nrows=nrows)


@st.cache_data(show_spinner=False)
def read_excel_sheets(file_bytes: bytes) -> list:
    from io import BytesIO
    xls = pd.ExcelFile(BytesIO(file_bytes))
    return xls.sheet_names


@st.cache_data(show_spinner=False)
def read_excel_preview(file_bytes: bytes, sheet_name: str, nrows: int = 15) -> pd.DataFrame:
    from io import BytesIO
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, nrows=nrows)


# =====================
# Sidebar — Upload & Save
# =====================
with st.sidebar:
    st.header("Upload & Save")
    uploaded_files = st.file_uploader(
        "Upload Excel/CSV files",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="Files will be saved under ./uploads with unique names.",
    )
    save_btn = st.button("Save uploaded files", use_container_width=True)

# Show saved files table (always visible)
st.subheader("Saved uploads")
st.dataframe(list_saved_files(UPLOAD_DIR), use_container_width=True, height=220)

# Handle saving
if save_btn:
    if not uploaded_files:
        st.warning("No files selected. Please upload at least one file.")
    else:
        saved_paths = []
        for uf in uploaded_files:
            try:
                p = save_uploaded_file(uf)
                saved_paths.append(p)
            except Exception as e:
                st.error(f"Failed to save **{uf.name}**: {e}")
        if saved_paths:
            st.success(f"Saved {len(saved_paths)} file(s) to ./uploads")
            st.toast("Uploads saved successfully", icon="✅")
            st.rerun()

# =====================
# Previews for current session uploads
# =====================
st.subheader("Preview of uploaded files (this session)")
if uploaded_files:
    for uf in uploaded_files:
        with st.expander(f"Preview: {uf.name}", expanded=False):
            try:
                file_bytes = uf.getvalue()
                if uf.name.lower().endswith(".csv"):
                    df_preview = read_csv_preview(file_bytes)
                    st.dataframe(df_preview, use_container_width=True)
                else:
                    # Excel: let user choose sheet
                    sheets = read_excel_sheets(file_bytes)
                    sheet = st.selectbox(
                        "Select sheet",
                        sheets,
                        key=f"sheet_{uf.name}_{len(sheets)}",
                    )
                    df_preview = read_excel_preview(file_bytes, sheet)
                    st.dataframe(df_preview, use_container_width=True)
            except Exception as e:
                st.error(f"Could not preview {uf.name}: {e}")
else:
    st.info("⬆️ Upload one or more files to see previews here.")


# =====================
# Department Report generation (auto-named)
# =====================
st.subheader("Generate Department Report with Charts (auto-named)")
st.caption("Creates an Excel file stored under ./reports and offers a download.")

# Your in-memory demo data (can later be replaced by parsed uploads)
data = {
    "Supply Chain": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Purchases Made": [25, 30, 28],
        "Procurement Plan Monitoring": [80, 85, 90],
        "Special Group Contracts": [5, 6, 4],
        "Suppliers Registered": [10, 15, 12],
    }),
    "Human Resources": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Staff Training": [2, 3, 4],
        "Staff Welfare Activities": [1, 2, 2],
        "Complaints Received": [3, 4, 2],
        "Complaints Resolved": [2, 4, 2],
        "Students on Industrial Training": [5, 6, 7],
    }),
    "Road Asset and Corridor Management": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Road Works Progress (%)": [60, 75, 85],
        "Inspections Done": [8, 10, 12],
        "Achievements Reported": [5, 6, 7],
    }),
    "Transport": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Service and Maintenance": [10, 12, 9],
        "Fuel Consumption (Litres)": [500, 600, 550],
    }),
    "Survey": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Surveys Completed": [10, 12, 14],
        "Pending Reports": [2, 1, 0],
    }),
    "Finance and Accounts": pd.DataFrame({
        "Contracts Paid": ["Contract A", "Contract B", "Contract C"],
        "Amount Paid": [10000, 15000, 12000],
        "Per Diem Paid": [3000, 2500, 2700],
        "Budget Consumption": [18000, 20000, 19000],
    }),
}

sheet_names_fixed = {
    "Supply Chain": "Supply Chain",
    "Human Resources": "HR",
    "Road Asset and Corridor Management": "Road Assets",
    "Transport": "Transport",
    "Survey": "Survey",
    "Finance and Accounts": "Finance",
}

col1, col2 = st.columns([1, 1])
with col1:
    base_name = st.text_input(
        "Base report filename (will get a timestamp)",
        value="Department_Report_with_Charts",
    )
with col2:
    gen_btn = st.button("Generate Report", type="primary", use_container_width=True)

if gen_btn:
    try:
        ts = time.strftime("%Y%m%d-%H%M%S")
        out_name = f"{base_name}_{ts}.xlsx"

        # Write to BytesIO for download and then persist the same bytes to disk
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            workbook = writer.book
            for dept, df in data.items():
                sheet_name = sheet_names_fixed[dept]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]

                # Only chart if "Month" present
                if "Month" in df.columns:
                    try:
                        chart = workbook.add_chart({"type": "line"})
                        for i, col in enumerate(df.columns[1:], start=1):
                            chart.add_series({
                                "name": [sheet_name, 0, i],
                                "categories": [sheet_name, 1, 0, len(df), 0],
                                "values": [sheet_name, 1, i, len(df), i],
                            })
                        chart.set_title({"name": f"{sheet_name} - Multi-Month Trend"})
                        chart.set_x_axis({"name": "Month"})
                        chart.set_y_axis({"name": "Value"})
                        worksheet.insert_chart("G2", chart)
                    except Exception as e:
                        st.warning(f"Chart skipped for {sheet_name}: {e}")

        buffer.seek(0)
        # Save to ./reports
        out_path = REPORT_DIR / out_name
        with open(out_path, "wb") as f:
            f.write(buffer.getvalue())

        st.success(f"Report generated: {out_name} (stored in ./reports)")
        st.download_button(
            label="⬇️ Download the Excel report",
            data=buffer,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Refresh reports table
        st.dataframe(list_saved_files(REPORT_DIR), use_container_width=True, height=220)

    except Exception as e:
        st.error(f"Failed to generate the report: {e}")


# =====================
# Download past reports
# =====================
st.subheader("Download previous reports")
rep_df = list_saved_files(REPORT_DIR)
if not rep_df.empty:
    options = rep_df.sort_values("modified", ascending=False)["file"].tolist()
    selected = st.selectbox("Select a report to download", options)
    if selected:
        rep_path = REPORT_DIR / selected
        with open(rep_path, "rb") as f:
            st.download_button(
                label=f"Download {selected}",
                data=f.read(),
                file_name=selected,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
else:
    st.info("No reports found yet. Generate one above.")


# =====================
# Notes
# =====================
st.caption(
    "ℹ️ Reminder: If deployed on Streamlit Community Cloud, local files are ephemeral. "
    "Use S3/Drive for permanent storage if needed."
)

