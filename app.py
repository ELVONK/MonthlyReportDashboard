# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import re, time, uuid, json
import shutil
import xlsxwriter
import plotly.express as px
import plotly.io as pio

# ---------------------
# Page config & theme CSS
# ---------------------
st.set_page_config(page_title="Monthly Report Dashboard", layout="wide")

# -- color palette (yellow, gold, green) applied via CSS --
_THEME_CSS = """
:root{
  --brand-yellow: #FFF8E1; /* very light yellow background */
  --brand-gold: #D4AF37;   /* gold for buttons and accents */
  --brand-green: #2E7D32;  /* green for success and highlights */
  --brand-dark: #263238;   /* dark text */
}

.reportview-container, .main, .block-container { background-color: var(--brand-yellow) !important; }
.stButton>button, .stDownloadButton>button { background: linear-gradient(180deg, var(--brand-gold), #b58f2c) !important; color: white !important; border-radius: 10px !important; }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #fffbe6, #f7f3e6) !important; }
.stAlert-success { background-color: rgba(46,125,50,0.08) !important; border-left: 4px solid var(--brand-green) !important; }
"""
st.markdown(f"<style>{_THEME_CSS}</style>", unsafe_allow_html=True)
st.markdown("<h1 style='color:#263238;margin-bottom:0.2rem'>Monthly Report Dashboard</h1>", unsafe_allow_html=True)
st.caption("Preview saved files, map columns to department sheets, and generate auto-named Excel reports.")

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
    missing_deps.append("xlrd (optional for .xls)")

# ---------------------
# Paths & helpers
# ---------------------
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
ARCHIVE_DIR = REPORT_DIR / "archive"
MAPPINGS_FILE = BASE_DIR / "mappings.json"
for d in (UPLOAD_DIR, REPORT_DIR, ARCHIVE_DIR):
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


def _choose_excel_engine_from_filename(fname: str) -> dict:
    fname = fname.lower()
    if fname.endswith('.xlsx'):
        if not HAS_OPENPYXL:
            raise ImportError('openpyxl not installed')
        return {'engine': 'openpyxl'}
    if fname.endswith('.xls'):
        if not HAS_XLRD:
            raise ImportError('xlrd not installed')
        return {'engine': 'xlrd'}
    return {}


def read_excel_sheets_from_bytes(file_bytes: bytes, fname: str) -> list:
    engine_kw = _choose_excel_engine_from_filename(fname)
    return pd.ExcelFile(BytesIO(file_bytes), **engine_kw).sheet_names


def read_excel_preview_from_bytes(file_bytes: bytes, sheet_name: str, fname: str, nrows: int = 15) -> pd.DataFrame:
    engine_kw = _choose_excel_engine_from_filename(fname)
    return pd.read_excel(BytesIO(file_bytes), sheet_name=sheet_name, nrows=nrows, **engine_kw)


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
# Predefined departments
# ---------------------
DEPARTMENTS = ["Supply Chain","Human Resources","Road Assets","Transport","Survey","Finance"]

# Color palette for charts
PALETTE = ["#D4AF37", "#2E7D32", "#263238", "#F6E27A", "#9CCC65"]

# ---------------------
# Sidebar
# ---------------------
with st.sidebar:
    st.header("Files & Actions")
    uploaded_files = st.file_uploader("Upload Excel/CSV files", type=["xlsx", "xls", "csv"], accept_multiple_files=True)
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
        chosen_for_report = []
        chosen_for_mapping = ""
    else:
        saved_options = saved_df.sort_values("modified", ascending=False)["file"].tolist()
        chosen_for_report = st.multiselect("Files to include in report", options=saved_options, default=None)
        chosen_for_mapping = st.selectbox("File to map/preview", options=[""] + saved_options)

    st.markdown("---")
    if st.button("Open Reports Folder (in app) "):
        st.info(f"Reports are stored under: {REPORT_DIR}")

# ---------------------
# Main layout
# ---------------------
mappings = load_mappings()

tabs = st.tabs(["Preview", "Mappings", "Generate Report", "Past Reports"])

# Preview tab
with tabs[0]:
    st.header("Preview files")
    st.markdown("Tip: preview session uploads or saved files.")
    if uploaded_files:
        for uf in uploaded_files:
            with st.expander(f"{uf.name}"):
                try:
                    fb = uf.getvalue()
                    if uf.name.lower().endswith('.csv'):
                        st.dataframe(pd.read_csv(BytesIO(fb), nrows=30), use_container_width=True)
                    else:
                        try:
                            engine_kw = _choose_excel_engine_from_filename(uf.name)
                            sheets = pd.ExcelFile(BytesIO(fb), **engine_kw).sheet_names
                            sheet = st.selectbox(f"Sheet ({uf.name})", sheets, key=f"sess_{uf.name}")
                            st.dataframe(pd.read_excel(BytesIO(fb), sheet_name=sheet, nrows=50, **engine_kw), use_container_width=True)
                        except Exception as e:
                            st.error(f"Preview failed: {e}")
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
                    else:
                        try:
                            engine_kw = _choose_excel_engine_from_filename(fname)
                            sheets = pd.ExcelFile(BytesIO(fb), **engine_kw).sheet_names
                            sheet = st.selectbox(f"Sheet ({fname})", sheets, key=f"saved_{fname}")
                            st.dataframe(pd.read_excel(BytesIO(fb), sheet_name=sheet, nrows=50, **engine_kw), use_container_width=True)
                        except Exception as e:
                            st.error(f"Preview failed: {e}")
                except Exception as e:
                    st.error(e)
    else:
        st.info("No saved uploads found. Save uploaded files first.")

# Mappings tab
with tabs[1]:
    st.header("Column -> Department Mapping")
    st.markdown("Create or edit mappings that tell the generator how to place columns into department sheets.")
    file_for_map = st.selectbox("Choose a saved file to create/edit mapping", options=[""] + saved_options)
    if file_for_map:
        p = UPLOAD_DIR / file_for_map
        try:
            fb = read_file_bytes_from_disk(p)
            if file_for_map.lower().endswith('.csv'):
                df_full = pd.read_csv(BytesIO(fb))
                sheet_list = ["(csv)"]
                chosen_sheet = sheet_list[0]
            else:
                engine_kw = _choose_excel_engine_from_filename(file_for_map)
                xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                sheet_list = xls.sheet_names
                chosen_sheet = st.selectbox("Select sheet to inspect for columns", sheet_list)
                df_full = pd.read_excel(BytesIO(fb), sheet_name=chosen_sheet, **engine_kw)
            if not df_full.empty:
                cols = list(df_full.columns)
                st.write(cols)
                map_key = f"{file_for_map}::{chosen_sheet}"
                current_map = mappings.get(map_key, {})
                new_map = {}
                for c in cols:
                    default = current_map.get(c, "")
                    try:
                        idx = 0 if default == "" else (DEPARTMENTS.index(default) + 1)
                    except Exception:
                        idx = 0
                    target = st.selectbox(f"Column: {c}", options=["", *DEPARTMENTS], index=idx, key=f"map_{map_key}_{c}")
                    if target:
                        new_map[c] = target
                if st.button("Save mapping for this sheet"):
                    mappings[map_key] = new_map
                    save_mappings(mappings)
                    st.success("Mapping saved")
                    st.experimental_rerun()
                if current_map:
                    st.json(current_map)
        except Exception as e:
            st.error(f"Failed to open file for mapping: {e}")
    else:
        st.info("Select a saved file to create or edit mappings.")

# Generate Report tab (with styling options and help)
with tabs[2]:
    st.header("Generate report")
    st.markdown("Select saved files (in the sidebar) to include in the report, or leave empty to use demo data.")

    base_name = st.text_input("Base report filename", value="Department_Report_with_Charts")
    generate = st.button("Generate Report (use selected files)")

    # Help / Examples
    with st.expander("Help & chart guidance", expanded=False):
        st.markdown("""
- **Line chart**: good for trends over time (e.g., Month on X axis, numeric on Y).
- **Bar chart**: good for comparing categories; enable *Stacked* for stacked bars.
- **Pie chart**: best for showing parts of a whole; choose a single value column and a category column for labels.

Tips:
- Ensure your X column is categorical or datetime.
- Numeric series should contain numbers (or be convertible). The app will coerce where possible.
- Use the HTML download to open charts in fullscreen and print from the browser.
""")

    st.markdown("---")
    st.subheader("Chart options (choose sheet, chart type, X axis, series, row range, styling)")

    def gather_available_sheets(selected_files: list) -> dict:
        out = {}
        for fname in (selected_files or []):
            p = UPLOAD_DIR / fname
            if not p.exists():
                continue
            try:
                fb = read_file_bytes_from_disk(p)
                if fname.lower().endswith('.csv'):
                    out[fname] = {'(csv)': pd.read_csv(BytesIO(fb))}
                else:
                    try:
                        engine_kw = _choose_excel_engine_from_filename(fname)
                    except ImportError:
                        continue
                    xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                    out[fname] = {}
                    for s in xls.sheet_names:
                        out[fname][s] = pd.read_excel(BytesIO(fb), sheet_name=s, **engine_kw)
            except Exception:
                continue
        return out

    available_sheets = gather_available_sheets(chosen_for_report)
    sheet_choices = []
    sheet_map = {}
    for fname, sheets in available_sheets.items():
        for sname in sheets.keys():
            key = f"{fname}::{sname}"
            sheet_choices.append(key)
            sheet_map[key] = (fname, sname)

    selected_sheet_for_chart = st.selectbox("Select sheet to configure chart", options=[""] + sheet_choices)

    chart_config = {}
    if selected_sheet_for_chart:
        fname, sname = sheet_map[selected_sheet_for_chart]
        df = available_sheets[fname][sname]
        st.markdown(f"**Preview of {fname} / {sname} (top 10 rows)**")
        st.dataframe(df.head(10), use_container_width=True)

        cols = list(df.columns)
        chart_type = st.selectbox("Chart type", options=["Line", "Bar", "Pie"], index=0)
        x_col = st.selectbox("Choose X-axis column (typically Month or Category)", options=[""] + cols)
        if chart_type == 'Pie':
            y_cols = st.selectbox("Choose value column (single) for Pie chart", options=[""] + cols)
            y_cols = [y_cols] if y_cols else []
        else:
            y_cols = st.multiselect("Choose one or more series columns (numeric)", options=cols)

        # Styling options
        st.markdown("**Styling**")
        color_scheme = st.selectbox("Color scheme", options=["Gold & Green", "Default"], index=0)
        if color_scheme == "Gold & Green":
            colors = PALETTE
        else:
            colors = None

        if chart_type == 'Bar':
            stacked = st.checkbox("Stack bars", value=False)
        else:
            stacked = False
        if chart_type == 'Line':
            smoothing = st.checkbox("Smooth line (spline)", value=False)
            show_markers = st.checkbox("Show markers", value=True)
        else:
            smoothing = False
            show_markers = True

        # Robust slider
        try:
            max_rows = int(max(1, len(df)))
        except Exception:
            max_rows = 1
        default_low, default_high = 1, max_rows
        if default_high < default_low:
            default_high = default_low
        default_low, default_high = int(default_low), int(default_high)
        try:
            r1, r2 = st.slider("Row range (1-indexed)", min_value=1, max_value=max_rows, value=(default_low, default_high), step=1)
        except Exception as slider_err:
            st.warning(f"Could not show row-range slider (using full range). Details: {slider_err}")
            r1, r2 = 1, max_rows

        chart_config = {
            'fname': fname,
            'sname': sname,
            'chart_type': chart_type,
            'x_col': x_col,
            'y_cols': y_cols,
            'row_range': (r1 - 1, r2 - 1),
            'colors': colors,
            'stacked': stacked,
            'smoothing': smoothing,
            'show_markers': show_markers,
        }

        # Chart preview using Plotly
        if x_col and y_cols:
            try:
                preview_df = df.iloc[chart_config['row_range'][0]:chart_config['row_range'][1] + 1].copy()
                for yc in chart_config['y_cols']:
                    preview_df[yc] = pd.to_numeric(preview_df[yc], errors='coerce')

                line_shape = 'spline' if chart_config['smoothing'] else 'linear'
                if chart_config['chart_type'] == 'Pie':
                    fig = px.pie(preview_df, names=chart_config['x_col'], values=chart_config['y_cols'][0], title=f"Preview: {fname} / {sname} (Pie)", color_discrete_sequence=chart_config['colors'])
                elif chart_config['chart_type'] == 'Bar':
                    fig = px.bar(preview_df, x=chart_config['x_col'], y=chart_config['y_cols'], title=f"Preview: {fname} / {sname} (Bar)", color_discrete_sequence=chart_config['colors'])
                    if chart_config['stacked']:
                        fig.update_layout(barmode='stack')
                else:
                    fig = px.line(preview_df, x=chart_config['x_col'], y=chart_config['y_cols'], title=f"Preview: {fname} / {sname} (Line)", line_shape=line_shape, markers=chart_config['show_markers'], color_discrete_sequence=chart_config['colors'])

                st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': True, 'scrollZoom': True})
                html = pio.to_html(fig, full_html=True)
                st.download_button(label="Download chart as HTML (open to view fullscreen/print)", data=html, file_name=f"{Path(fname).stem}_{sname}_chart.html", mime='text/html')
                try:
                    img_bytes = fig.to_image(format='png')
                    st.download_button(label="Download chart as PNG", data=img_bytes, file_name=f"{Path(fname).stem}_{sname}_chart.png", mime='image/png')
                except Exception:
                    st.caption("PNG export unavailable (install 'kaleido' to enable PNG export).")
            except Exception as e:
                st.warning(f"Could not render chart preview: {e}")
        else:
            st.info("Choose an X column and at least one series column/value to preview the chart.")

    def gather_data_for_report(selected_files: list) -> dict:
        if not selected_files:
            return {"Supply Chain": pd.DataFrame({"Month":["Jan","Feb","Mar"], "Purchases Made":[25,30,28]}), "Human Resources": pd.DataFrame({"Month":["Jan","Feb","Mar"], "Staff Training":[2,3,4]})}
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
                else:
                    engine_kw = _choose_excel_engine_from_filename(fname)
                    xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                    for s in xls.sheet_names:
                        df = pd.read_excel(BytesIO(fb), sheet_name=s, **engine_kw)
                        out[f"{Path(fname).stem[:20]}_{s[:8]}"] = df
            except Exception as e:
                st.warning(f"Failed to read {fname}: {e}")
        return out

    if generate:
        try:
            chosen = chosen_for_report
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
                        safe_name = re.sub(r'[:\/?*\[\]]', "_", str(sheet_name))[:31]
                        assigned = False
                        for k, mp in mappings.items():
                            fname_key = k.split('::')[0]
                            if fname_key in sheet_name:
                                dept_groups = {}
                                for col, dep in mp.items():
                                    if col in df.columns:
                                        dept_groups.setdefault(dep, []).append(col)
                                for dep, cols in dept_groups.items():
                                    subdf = df[cols]
                                    sub_name = (dep[:22] + "_" + safe_name[:6])[:31]
                                    subdf.to_excel(writer, sheet_name=sub_name, index=False)
                                assigned = True
                                break
                        if not assigned:
                            df.to_excel(writer, sheet_name=safe_name, index=False)
                        # insert chart if chart_config matches this sheet
                        if chart_config and (chart_config['fname'] in sheet_name or chart_config['sname'] in sheet_name):
                            try:
                                x = chart_config['x_col']
                                ys = chart_config['y_cols']
                                r0, r1 = chart_config['row_range']
                                ctype = chart_config.get('chart_type', 'Line')
                                worksheet = writer.sheets.get(safe_name)
                                if worksheet is None:
                                    worksheet = writer.book.add_worksheet(safe_name)
                                if ctype in ('Line', 'Bar') and ys:
                                    chart = workbook.add_chart({'type': 'line' if ctype == 'Line' else 'column'})
                                    for col in ys:
                                        try:
                                            col_idx = list(df.columns).index(col)
                                            chart.add_series({
                                                'name': [safe_name, 0, col_idx],
                                                'categories': [safe_name, 1 + r0, list(df.columns).index(x), 1 + r1, list(df.columns).index(x)],
                                                'values': [safe_name, 1 + r0, col_idx, 1 + r1, col_idx],
                                            })
                                        except ValueError:
                                            continue
                                    chart.set_title({'name': f"{safe_name} - {ctype} Chart"})
                                    chart.set_x_axis({'name': x})
                                    worksheet.insert_chart('G2', chart)
                                elif ctype == 'Pie' and ys:
                                    val_col = ys[0]
                                    try:
                                        val_idx = list(df.columns).index(val_col)
                                        cat_idx = list(df.columns).index(x)
                                        chart = workbook.add_chart({'type': 'pie'})
                                        chart.add_series({
                                            'name': val_col,
                                            'categories': [safe_name, 1 + r0, cat_idx, 1 + r1, cat_idx],
                                            'values': [safe_name, 1 + r0, val_idx, 1 + r1, val_idx],
                                        })
                                        chart.set_title({'name': f"{safe_name} - Pie Chart"})
                                        worksheet.insert_chart('G2', chart)
                                    except Exception:
                                        pass
                            except Exception as e:
                                st.warning(f"Failed to insert custom chart into Excel: {e}")
                buffer.seek(0)
                out_path = REPORT_DIR / out_name
                with open(out_path, 'wb') as f:
                    f.write(buffer.getvalue())
                st.success(f"Report generated: {out_name}")
                st.download_button("Download report", data=buffer, file_name=out_name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                st.dataframe(list_saved_files(REPORT_DIR), use_container_width=True)
        except Exception as e:
            st.error(f"Report generation failed: {e}")

# Past Reports tab
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
        st.markdown('---')
        st.subheader('Manage reports')
        manage_choices = rep_df.sort_values('modified', ascending=False)['file'].tolist()
        to_manage = st.multiselect('Select report(s) to archive or delete', options=manage_choices)
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button('Archive selected'):
                if not to_manage:
                    st.warning('No reports selected')
                else:
                    for fn in to_manage:
                        src = REPORT_DIR / fn
                        dst = ARCHIVE_DIR / fn
                        try:
                            shutil.move(str(src), str(dst))
                        except Exception as e:
                            st.error(f'Failed to archive {fn}: {e}')
                    st.success('Selected reports archived')
                    st.experimental_rerun()
        with col_b:
            if st.button('Delete selected'):
                if not to_manage:
                    st.warning('No reports selected')
                else:
                    for fn in to_manage:
                        try:
                            (REPORT_DIR / fn).unlink()
                        except Exception as e:
                            st.error(f'Failed to delete {fn}: {e}')
                    st.success('Selected reports deleted')
                    st.experimental_rerun()

# End of file
