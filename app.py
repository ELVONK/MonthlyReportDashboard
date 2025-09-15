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
