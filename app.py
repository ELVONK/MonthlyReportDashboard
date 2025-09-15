# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import re, time, uuid, json
import xlsxwriter

st.set_page_config(page_title="Monthly Report Dashboard", layout="wide")

# ---------------------
# Dependency checks
# ---------------------
missing_deps = []
try:
    import openpyxl  # noqa: F401
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False
    missing_deps.append("openpyxl")

try:
    import xlrd  # noqa: F401
    HAS_XLRD = True
except Exception:
    HAS_XLRD = False
    # xlrd is only required for older .xls — pandas may also use other engines
    missing_deps.append("xlrd (optional for .xls)")

# ---------------------
# Paths & helpers
# ---------------------
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
MAPPINGS_FILE = BASE_DIR / "mappings.json"
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
        raise ImportError("openpyxl not installed")
    return pd.ExcelFile(BytesIO(file_bytes), engine="openpyxl").sheet_names


def read_excel_preview_from_bytes(file_bytes: bytes, sheet_name: str, nrows: int = 15) -> pd.DataFrame:
    if not HAS_OPENPYXL:
        raise ImportError("openpyxl not installed")
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, nrows=nrows, engine="openpyxl")


def load_mappings() -> dict:
    if MAPPINGS_FILE.exists():
        try:
            return json.loads(MAPPINGS_FILE.read_text())
        except Exception:
            return {}
    return {}


def save_mappings(d: dict):
    MAPPINGS_FILE.write_text(json.dumps(d, indent=2))


# ---------------------
# Predefined departments & UI utility
# ---------------------
DEPARTMENTS = [
    "Supply Chain",
    "Human Resources",
    "Road Assets",
    "Transport",
    "Survey",
    "Finance",
]


# ---------------------
# Top-level UI: dependency banner + instructions
# ---------------------
with st.container():
    col1, col2 = st.columns([6, 1])
    with col1:
        st.title("Monthly Report Dashboard")
        st.caption("Preview saved files, map columns to department sheets, and generate auto-named Excel reports.")
    with col2:
        if missing_deps:
            st.error("Missing optional deps: " + ", ".join(missing_deps))
            with st.expander("How to install missing dependencies"):
                st.markdown("""
- For `.xlsx` files: `pip install openpyxl`
- For older `.xls` files: `pip install xlrd`

Add these to your `requirements.txt` before deploying.
""")
        else:
            st.success("All optional dependencies available")

st.markdown("---")


# ---------------------
# Sidebar: Upload, saved files and quick actions
# ---------------------
with st.sidebar:
    st.header("Files & Actions")
    uploaded_files = st.file_uploader(
        "Upload Excel/CSV files",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="Upload then Save to persist under ./uploads.",
    )
    if st.button("Save uploaded files", use_container_width=True):
        if not uploaded_files:
            st.warning("Choose files first")
        else:
            saved = []
            for uf in uploaded_files:
                try:
                    saved.append(save_uploaded_file(uf))
                except Exception as e:
                    st.error(f"Failed to save {uf.name}: {e}")
            if saved:
                st.success(f"Saved {len(saved)} file(s)")
                st.experimental_rerun()

    st.markdown("---")
    st.subheader("Saved uploads")
    saved_df = list_saved_files(UPLOAD_DIR)
    if saved_df.empty:
        st.info("No saved files yet.")
        saved_options = []
    else:
        saved_options = saved_df.sort_values("modified", ascending=False)["file"].tolist()
        # allow selecting multiple files for report generation
        chosen_for_report = st.multiselect("Files to include in report", options=saved_options, default=None)
        # single select for mapping & preview convenience
        chosen_for_mapping = st.selectbox("File to map/preview", options=[""] + saved_options)

    st.markdown("---")
    st.subheader("Quick actions")
    if st.button("Open Reports Folder (in app) "):
        st.info(f"Reports are stored under: {REPORT_DIR}")


# ---------------------
# Main area: tabs for Preview, Mappings, Generate, Reports
# ---------------------
mappings = load_mappings()

tabs = st.tabs(["Preview", "Mappings", "Generate Report", "Past Reports"])

# ----- Tab: Preview -----
with tabs[0]:
    st.header("Preview files")
    st.markdown("Tip: You can preview session uploads (top) or saved files (bottom).")

    st.subheader("Session uploads")
    if uploaded_files:
        for uf in uploaded_files:
            with st.expander(f"{uf.name}"):
                try:
                    fb = uf.getvalue()
                    if uf.name.lower().endswith('.csv'):
                        st.dataframe(pd.read_csv(BytesIO(fb), nrows=30), use_container_width=True)
                    elif uf.name.lower().endswith('.xls'):
                        try:
                            st.dataframe(pd.read_excel(BytesIO(fb), sheet_name=0, nrows=30), use_container_width=True)
                        except Exception as e:
                            st.error(f"Preview failed (.xls): {e}")
                    else:
                        if not HAS_OPENPYXL:
                            st.error("Cannot preview .xlsx: openpyxl not installed")
                        else:
                            sheets = read_excel_sheets_from_bytes(fb)
                            sheet = st.selectbox(f"Sheet ({uf.name})", sheets, key=f"sess_{uf.name}")
                            st.dataframe(read_excel_preview_from_bytes(fb, sheet, nrows=50), use_container_width=True)
                except Exception as e:
                    st.error(e)
    else:
        st.info("No session uploads. Use the sidebar to upload and save files.")

    st.markdown("---")
    st.subheader("Saved uploads")
    if saved_options:
        for fname in saved_options:
            with st.expander(fname):
                p = UPLOAD_DIR / fname
                try:
                    fb = read_file_bytes_from_disk(p)
                    if fname.lower().endswith('.csv'):
                        st.dataframe(pd.read_csv(BytesIO(fb), nrows=50), use_container_width=True)
                    elif fname.lower().endswith('.xls'):
                        try:
                            st.dataframe(pd.read_excel(BytesIO(fb), sheet_name=0, nrows=50), use_container_width=True)
                        except Exception as e:
                            st.error(f"Preview failed (.xls): {e}")
                    else:
                        if not HAS_OPENPYXL:
                            st.error("Cannot preview .xlsx: openpyxl not installed")
                        else:
                            sheets = read_excel_sheets_from_bytes(fb)
                            sheet = st.selectbox(f"Sheet ({fname})", sheets, key=f"saved_{fname}")
                            st.dataframe(read_excel_preview_from_bytes(fb, sheet, nrows=50), use_container_width=True)
                except Exception as e:
                    st.error(f"Could not read {fname}: {e}")
    else:
        st.info("No saved uploads found. Save uploaded files first.")

# ----- Tab: Mappings -----
with tabs[1]:
    st.header("Column -> Department Mapping")
    st.markdown("Create or edit mappings that tell the generator how to place columns into department sheets.")

    # choose a saved file to create mapping from
    file_for_map = st.selectbox("Choose a saved file to create/edit mapping", options=[""] + saved_options)
    if file_for_map:
        p = UPLOAD_DIR / file_for_map
        try:
            fb = read_file_bytes_from_disk(p)
            # choose sheet for xlsx/xls; csv uses single sheet
            if file_for_map.lower().endswith('.csv'):
                df_full = pd.read_csv(BytesIO(fb))
                sheet_list = ["(csv)"]
            else:
                if not HAS_OPENPYXL and file_for_map.lower().endswith('.xlsx'):
                    st.error("openpyxl required to edit mapping for .xlsx")
                    df_full = pd.DataFrame()
                    sheet_list = []
                else:
                    # try to read first sheet for mapping context, but allow selecting others
                    xls = pd.ExcelFile(BytesIO(fb), engine=("openpyxl" if file_for_map.lower().endswith('.xlsx') else None))
                    sheet_list = xls.sheet_names
                    chosen_sheet = st.selectbox("Select sheet to inspect for columns", sheet_list)
                    df_full = pd.read_excel(BytesIO(fb), sheet_name=chosen_sheet, engine=("openpyxl" if file_for_map.lower().endswith('.xlsx') else None))

            if not df_full.empty:
                st.markdown("**Columns detected:**")
                cols = list(df_full.columns)
                st.write(cols)

                # load existing mapping for this file (by filename + sheet)
                map_key = f"{file_for_map}::{sheet_list[0] if len(sheet_list)==1 else chosen_sheet}"
                current_map = mappings.get(map_key, {})

                st.markdown("---")
                st.markdown("**Assign columns to departments**")
                new_map = {}
                for c in cols:
                    # default: existing mapping or blank
                    default = current_map.get(c, "")
                    target = st.selectbox(f"Column: {c}", options=["", *DEPARTMENTS], index=(0 if default=="" else DEPARTMENTS.index(default)+1 if default in DEPARTMENTS else 0), key=f"map_{map_key}_{c}")
                    if target:
                        new_map[c] = target

                if st.button("Save mapping for this sheet"):
                    mappings[map_key] = new_map
                    save_mappings(mappings)
                    st.success("Mapping saved")
                    st.experimental_rerun()

                if current_map:
                    st.markdown("**Current mapping preview**")
                    st.json(current_map)

        except Exception as e:
            st.error(f"Failed to open file for mapping: {e}")
    else:
        st.info("Select a saved file to create or edit mappings.")

# ----- Tab: Generate Report -----
with tabs[2]:
    st.header("Generate report")
    st.markdown("Select saved files (in the sidebar) to include in the report, or leave empty to use demo data.")

    base_name = st.text_input("Base report filename", value="Department_Report_with_Charts")
    generate = st.button("Generate Report (use selected files)")

    def gather_data_for_report(selected_files: list) -> dict:
        if not selected_files:
            # fallback demo
            return {
                "Supply Chain": pd.DataFrame({"Month":["Jan","Feb","Mar"], "Purchases Made":[25,30,28]}),
                "Human Resources": pd.DataFrame({"Month":["Jan","Feb","Mar"], "Staff Training":[2,3,4]}),
            }
        out = {}
        for fname in selected_files:
            p = UPLOAD_DIR / fname
            if not p.exists():
                st.warning(f"Skipping missing {fname}")
                continue
            try:
                fb = read_file_bytes_from_disk(p)
                if fname.lower().endswith('.csv'):
                    df = pd.read_csv(BytesIO(fb))
                    out[fname[:25]] = df
                elif fname.lower().endswith('.xls'):
                    xls = pd.ExcelFile(BytesIO(fb))
                    for s in xls.sheet_names:
                        df = pd.read_excel(BytesIO(fb), sheet_name=s)
                        out[f"{Path(fname).stem[:20]}_{s[:8]}"] = df
                else:  # xlsx
                    if not HAS_OPENPYXL:
                        st.error(f"Cannot read {fname}: openpyxl missing")
                        continue
                    xls = pd.ExcelFile(BytesIO(fb), engine="openpyxl")
                    for s in xls.sheet_names:
                        df = pd.read_excel(BytesIO(fb), sheet_name=s, engine="openpyxl")
                        out[f"{Path(fname).stem[:20]}_{s[:8]}"] = df
            except Exception as e:
                st.warning(f"Failed to read {fname}: {e}")
        return out

    if generate:
        try:
            # read chosen files from sidebar variable `chosen_for_report` (may not exist if none)
            try:
                chosen = st.session_state.get('multiselect', None)
            except Exception:
                chosen = None
            # better: read files directly from the sidebar variable if present
            try:
                chosen = chosen_for_report  # from sidebar scope
            except Exception:
                chosen = None

            data_for_report = gather_data_for_report(chosen)

            if not data_for_report:
                st.error("No data available to include in report.")
            else:
                ts = time.strftime("%Y%m%d-%H%M%S")
                out_name = f"{base_name}_{ts}.xlsx"
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    for sheet_name, df in data_for_report.items():
                        safe_name = re.sub(r'[:\\/?*\[\]]', "_", str(sheet_name))[:31]
                        # apply mapping: if mappings exist for this originating file/sheet, group columns into department sheets
                        # simple approach: if any mapping matches this sheet's key, build per-department dfs
                        assigned = False
                        # find mapping keys that reference this source by filename prefix
                        for k, mp in mappings.items():
                            # k is file::sheet
                            if k.split('::')[0].startswith(sheet_name.split('_')[0]):
                                # apply mapping
                                dept_groups = {}
                                for col, dep in mp.items():
                                    if col in df.columns:
                                        dept_groups.setdefault(dep, []).append(col)
                                # write each department frame
                                for dep, cols in dept_groups.items():
                                    subdf = df[cols]
                                    writer.sheets  # ensure writer initialized
                                    sub_name = (dep[:22] + "_" + safe_name[:6])[:31]
                                    subdf.to_excel(writer, sheet_name=sub_name, index=False)
                                assigned = True
                                break

                        if not assigned:
                            # default behavior: write the dataframe as-is
                            df.to_excel(writer, sheet_name=safe_name, index=False)

                        # add simple chart if Month present
                        if "Month" in df.columns and df.shape[0] > 0 and df.shape[1] > 1:
                            try:
                                worksheet = writer.sheets.get(safe_name) or writer.book.add_worksheet(safe_name)
                                chart = workbook.add_chart({"type": "line"})
                                for i, col in enumerate(df.columns[1:], start=1):
                                    chart.add_series({
                                        "name": [safe_name, 0, i],
                                        "categories": [safe_name, 1, 0, len(df), 0],
                                        "values": [safe_name, 1, i, len(df), i],
                                    })
                                worksheet.insert_chart('G2', chart)
                            except Exception:
                                pass

                buffer.seek(0)
                out_path = REPORT_DIR / out_name
                with open(out_path, 'wb') as f:
                    f.write(buffer.getvalue())

                st.success(f"Report generated: {out_name}")
                st.download_button("Download report", data=buffer, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(list_saved_files(REPORT_DIR), use_container_width=True)

        except Exception as e:
            st.error(f"Report generation failed: {e}")

# ----- Tab: Past Reports -----
with tabs[3]:
    st.header("Past reports")
    rep_df = list_saved_files(REPORT_DIR)
    if rep_df.empty:
        st.info("No reports yet.")
    else:
        st.dataframe(rep_df.sort_values('modified', ascending=False), use_container_width=True)
        pick = st.selectbox("Download report", options=rep_df.sort_values('modified', ascending=False)['file'].tolist())
        if pick:
            p = REPORT_DIR / pick
            if p.exists():
                with open(p, 'rb') as f:
                    st.download_button(f"Download {pick}", data=f.read(), file_name=pick, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# End of file
