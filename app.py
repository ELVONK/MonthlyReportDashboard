# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import re, time, uuid
import xlsxwriter

st.set_page_config(page_title="Monthly Report Dashboard", layout="wide")
st.title("Monthly Report Dashboard — Upload, Preview & Auto-Named Reports")

# Check for optional dependency openpyxl early so we can show a friendly message
try:
    import openpyxl  # noqa: F401
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False
    _OPENPYXL_MSG = (
        "Missing optional dependency `openpyxl` required to read .xlsx files.\n\n"
        "Install it with:\n\n"
        "`pip install openpyxl`\n\n"
        "Or add `openpyxl` to your requirements.txt for deployments."
    )

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


def read_file_bytes_from_disk(path: Path) -> bytes:
    with open(path, "rb") as f:
        return f.read()


def read_excel_sheets_from_bytes(file_bytes: bytes) -> list:
    if not HAS_OPENPYXL:
        raise ImportError(_OPENPYXL_MSG)
    from io import BytesIO
    xls = pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl")
    return xls.sheet_names


def read_excel_preview_from_bytes(file_bytes: bytes, sheet_name: str, nrows: int = 15) -> pd.DataFrame:
    if not HAS_OPENPYXL:
        raise ImportError(_OPENPYXL_MSG)
    from io import BytesIO
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, nrows=nrows, engine="openpyxl")


# =====================
# Sidebar — Upload & Save & Saved-file controls
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

    st.markdown("---")
    st.header("Saved uploads")
    saved_df = list_saved_files(UPLOAD_DIR)
    if not saved_df.empty:
        # multi-select to pick files to include when generating reports
        saved_options = saved_df.sort_values("modified", ascending=False)["file"].tolist()
        selected_saved = st.multiselect(
            "Select saved file(s) to preview / include in generated report",
            options=saved_options,
            default=None,
            help="If you select saved file(s), Generate Report will use them instead of demo data.",
        )
        st.caption("Tip: Select multiple files to include each as separate sheet(s) in the generated report.")
    else:
        st.info("No saved uploads yet. Upload and Save files to see them here.")
        selected_saved = []

# Show saved files table (always visible in main area)
st.subheader("Saved uploads")
st.dataframe(saved_df, use_container_width=True, height=220)

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
            try:
                st.toast("Uploads saved successfully", icon="✅")
            except Exception:
                pass
            st.experimental_rerun()

# =====================
# Preview of current session uploads
# =====================
st.subheader("Preview of uploaded files (this session)")
if uploaded_files:
    for uf in uploaded_files:
        with st.expander(f"Preview (session): {uf.name}", expanded=False):
            try:
                file_bytes = uf.getvalue()
                fname = uf.name.lower()
                if fname.endswith(".csv"):
                    df_preview = pd.read_csv(BytesIO(file_bytes), nrows=15)
                    st.dataframe(df_preview, use_container_width=True)
                elif fname.endswith(".xls"):
                    try:
                        df_preview = pd.read_excel(BytesIO(file_bytes), sheet_name=0)
                        st.dataframe(df_preview, use_container_width=True)
                    except Exception as e:
                        st.error(f"Could not preview {uf.name}. For .xls files, ensure 'xlrd' is installed. Error: {e}")
                else:  # .xlsx
                    try:
                        sheets = read_excel_sheets_from_bytes(file_bytes)
                        sheet = st.selectbox("Select sheet", sheets, key=f"session_sheet_{uf.name}_{len(sheets)}")
                        df_preview = read_excel_preview_from_bytes(file_bytes, sheet)
                        st.dataframe(df_preview, use_container_width=True)
                    except ImportError:
                        st.error(
                            "❌ Cannot preview this .xlsx file because `openpyxl` is not installed. "
                            "Install it using `pip install openpyxl`."
                        )
                    except Exception as e:
                        st.error(f"Could not preview {uf.name}: {e}")
            except Exception as e:
                st.error(f"Could not process {uf.name}: {e}")
else:
    st.info("⬆️ Upload one or more files to see previews here.")


# =====================
# Preview saved uploads from disk
# =====================
st.subheader("Preview saved uploads (from ./uploads)")
if selected_saved:
    for fname in selected_saved:
        p = UPLOAD_DIR / fname
        if not p.exists():
            st.error(f"Saved file not found: {fname}")
            continue
        with st.expander(f"Preview (saved): {fname}", expanded=False):
            try:
                file_bytes = read_file_bytes_from_disk(p)
                if fname.lower().endswith(".csv"):
                    df_preview = pd.read_csv(BytesIO(file_bytes), nrows=15)
                    st.dataframe(df_preview, use_container_width=True)
                elif fname.lower().endswith(".xls"):
                    try:
                        # read first sheet
                        df_preview = pd.read_excel(BytesIO(file_bytes), sheet_name=0)
                        st.dataframe(df_preview, use_container_width=True)
                    except Exception as e:
                        st.error(f"Could not preview {fname}. For .xls files, ensure 'xlrd' is installed. Error: {e}")
                else:  # .xlsx
                    try:
                        sheets = read_excel_sheets_from_bytes(file_bytes)
                        sheet = st.selectbox("Select sheet", sheets, key=f"saved_sheet_{fname}_{len(sheets)}")
                        df_preview = read_excel_preview_from_bytes(file_bytes, sheet)
                        st.dataframe(df_preview, use_container_width=True)
                    except ImportError:
                        st.error(
                            "❌ Cannot preview this .xlsx file because `openpyxl` is not installed. "
                            "Install it using `pip install openpyxl`."
                        )
                    except Exception as e:
                        st.error(f"Could not preview {fname}: {e}")
            except Exception as e:
                st.error(f"Could not read {fname}: {e}")
else:
    st.info("Select saved upload(s) in the sidebar to preview them here.")


# =====================
# Department Report generation (auto-named)
# =====================
st.subheader("Generate Department Report with Charts (auto-named)")
st.caption("Creates an Excel file stored under ./reports and offers a download.")

# Demo data (used when no saved files selected)
demo_data = {
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
}

# Options: either use selected saved files OR demo_data
col1, col2 = st.columns([1, 1])
with col1:
    base_name = st.text_input(
        "Base report filename (will get a timestamp)",
        value="Department_Report_with_Charts",
    )
with col2:
    gen_btn = st.button("Generate Report", type="primary", use_container_width=True)

def gather_data_for_report(selected_saved_files: list) -> dict:
    """
    If selected_saved_files is non-empty, read those files and return a dict
    mapping sheet_name -> DataFrame (each sheet becomes a sheet in output).
    Otherwise return demo_data.
    """
    if not selected_saved_files:
        return demo_data

    out = {}
    for fname in selected_saved_files:
        p = UPLOAD_DIR / fname
        if not p.exists():
            st.warning(f"Skipping missing file: {fname}")
            continue
        try:
            fb = read_file_bytes_from_disk(p)
            if fname.lower().endswith(".csv"):
                df = pd.read_csv(BytesIO(fb))
                sheet_name = Path(fname).stem[:31]
                out[sheet_name] = df
            elif fname.lower().endswith(".xls"):
                # read all sheets if possible
                try:
                    xls = pd.ExcelFile(BytesIO(fb))
                    for s in xls.sheet_names:
                        df = pd.read_excel(BytesIO(fb), sheet_name=s)
                        sheet_name = f"{Path(fname).stem[:20]}_{s[:10]}".replace(" ", "_")[:31]
                        out[sheet_name] = df
                except Exception as e:
                    st.warning(f"Could not read .xls file {fname}: {e}")
            else:  # .xlsx
                if not HAS_OPENPYXL:
                    st.error(f"Cannot read {fname}: openpyxl not installed.")
                    continue
                try:
                    xls = pd.ExcelFile(BytesIO(fb), engine="openpyxl")
                    for s in xls.sheet_names:
                        df = pd.read_excel(BytesIO(fb), sheet_name=s, engine="openpyxl")
                        sheet_name = f"{Path(fname).stem[:20]}_{s[:10]}".replace(" ", "_")[:31]
                        out[sheet_name] = df
                except Exception as e:
                    st.warning(f"Could not read .xlsx file {fname}: {e}")
        except Exception as e:
            st.warning(f"Failed to load {fname}: {e}")
    return out

if gen_btn:
    try:
        # Determine source data
        report_data = gather_data_for_report(selected_saved)

        if not report_data:
            st.error("No data found to include in report (selected saved files were empty or unreadable).")
        else:
            ts = time.strftime("%Y%m%d-%H%M%S")
            out_name = f"{base_name}_{ts}.xlsx"

            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                workbook = writer.book
                for sheet_name, df in report_data.items():
                    # sanitize sheet name to 31 chars max and avoid invalid chars
                    safe_name = re.sub(r'[:\\/?*\[\]]', "_", str(sheet_name))[:31]
                    # if df has too many columns or index issues, let pandas handle and surface errors
                    df.to_excel(writer, sheet_name=safe_name, index=False)
                    worksheet = writer.sheets[safe_name]

                    # Add simple chart if there's a 'Month' column
                    if "Month" in df.columns and len(df.columns) > 1 and df.shape[0] > 0:
                        try:
                            chart = workbook.add_chart({"type": "line"})
                            for i, col in enumerate(df.columns[1:], start=1):
                                chart.add_series({
                                    "name": [safe_name, 0, i],
                                    "categories": [safe_name, 1, 0, len(df), 0],
                                    "values": [safe_name, 1, i, len(df), i],
                                })
                            chart.set_title({"name": f"{safe_name} - Trend"})
                            chart.set_x_axis({"name": "Month"})
                            chart.set_y_axis({"name": "Value"})
                            worksheet.insert_chart("G2", chart)
                        except Exception as e:
                            st.warning(f"Chart skipped for {safe_name}: {e}")

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
