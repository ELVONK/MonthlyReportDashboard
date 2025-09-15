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
st.caption("Preview saved files, map columns to department sheets, and generate auto-named Excel reports with analysis and previews.")

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
# Predefined departments and palette
# ---------------------
DEPARTMENTS = ["Supply Chain","Human Resources","Road Assets","Transport","Survey","Finance"]
PALETTE = ["#D4AF37", "#2E7D32", "#263238", "#F6E27A", "#9CCC65"]

# Ensure session_state keys exist
if 'custom_config' not in st.session_state:
    st.session_state['custom_config'] = None

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
    else:
        saved_options = saved_df.sort_values("modified", ascending=False)["file"].tolist()

    # Use a named key so it's accessible from other parts of the app
    chosen_for_report = st.multiselect("Files to include in report", options=saved_options, default=None, key='chosen_for_report')
    chosen_for_mapping = st.selectbox("File to map/preview", options=[""] + saved_options, key='chosen_for_mapping')

    st.markdown("---")
    if st.button("Open Reports Folder (in app) "):
        st.info(f"Reports are stored under: {REPORT_DIR}")

# ---------------------
# Main layout (tabs)
# ---------------------
mappings = load_mappings()

tabs = st.tabs(["Preview", "Mappings", "Generate Report", "Past Reports"])

# ----- Preview tab: includes auto analysis per saved sheet -----
with tabs[0]:
    st.header("Preview files & Auto Analysis")
    st.markdown("Preview session uploads or saved files — the app will also run a quick analysis and draw auto-charts.")

    def quick_analysis_and_chart(df: pd.DataFrame, title: str):
        st.markdown(f"#### {title}")
        if df.empty:
            st.info("Empty sheet")
            return
        # numeric columns
        num_cols = df.select_dtypes(include=['number']).columns.tolist()
        cat_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

        # show sample and summary
        st.write("Data sample:")
        st.dataframe(df.head(8), use_container_width=True)
        if num_cols:
            st.write("Summary stats for numeric columns:")
            st.dataframe(df[num_cols].describe().transpose(), use_container_width=True)
        else:
            st.info("No numeric columns found for detailed stats.")

        # Auto-chart heuristics
        fig = None
        if 'Month' in df.columns and num_cols:
            try:
                plot_df = df[['Month'] + num_cols].copy()
                for c in num_cols:
                    plot_df[c] = pd.to_numeric(plot_df[c], errors='coerce')
                fig = px.line(plot_df, x='Month', y=num_cols, title=f"Auto: {title} — Multi-series over Month", markers=True, color_discrete_sequence=PALETTE)
            except Exception:
                fig = None
        elif num_cols and cat_cols:
            try:
                agg = df.groupby(cat_cols[0])[num_cols[0]].sum().reset_index()
                fig = px.bar(agg, x=cat_cols[0], y=num_cols[0], title=f"Auto: {title} — {num_cols[0]} by {cat_cols[0]}", color_discrete_sequence=PALETTE)
            except Exception:
                fig = None
        elif num_cols:
            try:
                plot_df = df[num_cols].copy()
                plot_df = plot_df.apply(pd.to_numeric, errors='coerce')
                fig = px.line(plot_df.reset_index(), x='index', y=num_cols[0], title=f"Auto: {title} — {num_cols[0]}", color_discrete_sequence=PALETTE)
            except Exception:
                fig = None

        if fig is not None:
            st.plotly_chart(fig, use_container_width=True, config={'displayModeBar': True})
            html = pio.to_html(fig, full_html=True)
            st.download_button(label='Download auto-chart as HTML', data=html, file_name=f"{re.sub('[^0-9a-zA-Z]+','_', title)}_auto_chart.html", mime='text/html')
        else:
            st.info('No suitable auto-chart could be generated for this sheet.')

    # session uploads preview and analysis
    if uploaded_files:
        for uf in uploaded_files:
            with st.expander(f"(session) {uf.name}"):
                try:
                    fb = uf.getvalue()
                    if uf.name.lower().endswith('.csv'):
                        df = pd.read_csv(BytesIO(fb))
                        quick_analysis_and_chart(df, f"{uf.name}")
                    else:
                        engine_kw = _choose_excel_engine_from_filename(uf.name)
                        xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                        for s in xls.sheet_names:
                            df = pd.read_excel(BytesIO(fb), sheet_name=s, **engine_kw)
                            quick_analysis_and_chart(df, f"{uf.name}::{s}")
                except Exception as e:
                    st.error(f"Preview failed: {e}")
    else:
        st.info("No session uploads. Use the sidebar to upload and save files.")

    st.markdown('---')
    st.subheader('Saved uploads analysis')
    saved_df = list_saved_files(UPLOAD_DIR)
    if not saved_df.empty:
        for fname in saved_df.sort_values('modified', ascending=False)['file'].tolist():
            with st.expander(fname):
                try:
                    fb = read_file_bytes_from_disk(UPLOAD_DIR / fname)
                    if fname.lower().endswith('.csv'):
                        df = pd.read_csv(BytesIO(fb))
                        quick_analysis_and_chart(df, fname)
                    else:
                        engine_kw = _choose_excel_engine_from_filename(fname)
                        xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                        for s in xls.sheet_names:
                            df = pd.read_excel(BytesIO(fb), sheet_name=s, **engine_kw)
                            quick_analysis_and_chart(df, f"{fname}::{s}")
                except Exception as e:
                    st.error(f"Could not read {fname}: {e}")
    else:
        st.info('No saved uploads found yet.')

# ----- Mappings tab -----
with tabs[1]:
    st.header("Column -> Department Mapping")
    st.markdown("Create or edit mappings that tell the generator how to place columns into department sheets.")
    saved_opts_df = list_saved_files(UPLOAD_DIR)
    saved_opts = saved_opts_df.sort_values('modified', ascending=False)['file'].tolist() if not saved_opts_df.empty else []
    file_for_map = st.selectbox("Choose a saved file to create/edit mapping", options=[""] + saved_opts, key='file_for_map')
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
                chosen_sheet = st.selectbox("Select sheet to inspect for columns", sheet_list, key=f"map_sheet_{file_for_map}")
                df_full = pd.read_excel(BytesIO(fb), sheet_name=chosen_sheet, **engine_kw)
            if not df_full.empty:
                cols = list(df_full.columns)
                st.write(cols)
                map_key = f"{file_for_map}::{chosen_sheet}"
                mappings = load_mappings()
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

# ----- Generate Report tab (merged auto & custom charts, analysis embedding) -----
with tabs[2]:
    st.header("Generate report — auto charts + custom chart options")
    st.markdown("You can either let the app auto-generate charts per sheet (recommended), or configure a custom chart for one sheet and embed it.")

    base_name = st.text_input("Base report filename", value="Department_Report_with_Charts")
    generate = st.button("Generate Report")

    # Chart customization UI (as before) but fixed: uses sidebar-chosen files and session_state to persist
    st.markdown('---')
    st.subheader('Optional: Custom chart for one sheet')

    # use chosen_for_report from sidebar (stored in session_state)
    chosen_for_report = st.session_state.get('chosen_for_report', [])

    available_sheets = {}
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
                    engine_kw = _choose_excel_engine_from_filename(fname)
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

    custom_sheet = st.selectbox('Custom chart - pick a sheet (optional)', options=[""] + sheet_choices, key='custom_sheet')

    # Show stored custom config if present
    if st.session_state.get('custom_config'):
        st.caption('Saved custom chart config found — it will be applied when generating the report.')
        st.json(st.session_state['custom_config'])

    custom_config = None
    if custom_sheet:
        fname, sname = sheet_map[custom_sheet]
        df = available_sheets[fname][sname]
        st.markdown(f"Preview of {fname} / {sname}")
        st.dataframe(df.head(10), use_container_width=True)
        cols = list(df.columns)
        # make keys unique per sheet to avoid widget collisions
        safe_key = re.sub(r"[^0-9a-zA-Z]+", "_", custom_sheet)
        chart_type = st.selectbox("Chart type", options=["Line", "Bar", "Pie"], index=0, key=f"chart_type_{safe_key}")
        x_col = st.selectbox("X column", options=[""] + cols, key=f"xcol_{safe_key}")
        if chart_type == 'Pie':
            y_col_single = st.selectbox("Value column (single)", options=[""] + cols, key=f"ycol_single_{safe_key}")
            y_cols = [y_col_single] if y_col_single else []
        else:
            y_cols = st.multiselect("Series columns", options=cols, key=f"ycols_{safe_key}")
        # styling
        color_scheme = st.selectbox("Color scheme", options=["Gold & Green", "Default"], index=0, key=f"color_{safe_key}")
        colors = PALETTE if color_scheme == 'Gold & Green' else None
        smoothing = st.checkbox('Smooth line ( spline )', value=False, key=f"smooth_{safe_key}") if chart_type == 'Line' else False
        stacked = st.checkbox('Stack bars', value=False, key=f"stack_{safe_key}") if chart_type == 'Bar' else False
        # robust slider
        try:
            max_rows = int(max(1, len(df)))
        except Exception:
            max_rows = 1
        try:
            r1, r2 = st.slider('Row range (1-indexed)', min_value=1, max_value=max_rows, value=(1, max_rows), step=1, key=f"rows_{safe_key}")
        except Exception:
            r1, r2 = 1, max_rows

        tmp_config = {'fname': fname, 'sname': sname, 'chart_type': chart_type, 'x_col': x_col, 'y_cols': y_cols, 'row_range': (r1 - 1, r2 - 1), 'colors': colors, 'smoothing': smoothing, 'stacked': stacked}

        # preview
        if x_col and y_cols:
            try:
                preview_df = df.iloc[tmp_config['row_range'][0]:tmp_config['row_range'][1] + 1].copy()
                for yc in tmp_config['y_cols']:
                    preview_df[yc] = pd.to_numeric(preview_df[yc], errors='coerce')
                if chart_type == 'Pie':
                    fig = px.pie(preview_df, names=tmp_config['x_col'], values=tmp_config['y_cols'][0], color_discrete_sequence=tmp_config['colors'])
                elif chart_type == 'Bar':
                    fig = px.bar(preview_df, x=tmp_config['x_col'], y=tmp_config['y_cols'], color_discrete_sequence=tmp_config['colors'])
                    if tmp_config['stacked']:
                        fig.update_layout(barmode='stack')
                else:
                    shape = 'spline' if tmp_config['smoothing'] else 'linear'
                    fig = px.line(preview_df, x=tmp_config['x_col'], y=tmp_config['y_cols'], line_shape=shape, markers=True, color_discrete_sequence=tmp_config['colors'])
                st.plotly_chart(fig, use_container_width=True)
                st.download_button('Download custom chart as HTML', data=pio.to_html(fig, full_html=True), file_name=f"{Path(fname).stem}_{sname}_custom.html", mime='text/html')
            except Exception as e:
                st.warning(f'Could not render custom preview: {e}')
        else:
            st.info('Select X and Y columns to preview the custom chart')

        # allow user to save custom config so it persists across reruns
        if st.button('Save custom chart configuration', key=f"savecfg_{safe_key}"):
            st.session_state['custom_config'] = tmp_config
            st.success('Custom chart configuration saved — it will be embedded when you generate the report.')
        custom_config = st.session_state.get('custom_config')

    else:
        custom_config = st.session_state.get('custom_config')

    # When generating: write all sheets and embed charts (apply saved custom_config if present)
    if generate:
        try:
            data_for_report = {}
            # gather data from chosen files
            for fname in (chosen_for_report or []):
                p = UPLOAD_DIR / fname
                if not p.exists():
                    st.warning(f"Skipping missing {fname}")
                    continue
                fb = read_file_bytes_from_disk(p)
                if fname.lower().endswith('.csv'):
                    data_for_report[fname] = {'(csv)': pd.read_csv(BytesIO(fb))}
                else:
                    engine_kw = _choose_excel_engine_from_filename(fname)
                    xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                    data_for_report[fname] = {}
                    for s in xls.sheet_names:
                        data_for_report[fname][s] = pd.read_excel(BytesIO(fb), sheet_name=s, **engine_kw)

            if not data_for_report:
                st.error('No data selected for report. Choose saved files in the sidebar.')
            else:
                ts = time.strftime('%Y%m%d-%H%M%S')
                out_name = f"{base_name}_{ts}.xlsx"
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    workbook = writer.book
                    for fname, sheets in data_for_report.items():
                        for sname, df in sheets.items():
                            safe_name = re.sub(r'[:\/?*\[\]]', '_', f"{Path(fname).stem}_{sname}")[:31]
                            # write data
                            try:
                                df.to_excel(writer, sheet_name=safe_name, index=False)
                            except Exception:
                                df.head(1000).to_excel(writer, sheet_name=safe_name, index=False)
                            worksheet = writer.sheets[safe_name]

                            # Determine whether to apply custom config
                            applied_custom = False
                            cfg = st.session_state.get('custom_config')
                            if cfg and cfg.get('fname') == fname and cfg.get('sname') == sname:
                                applied_custom = True

                            if applied_custom and cfg:
                                try:
                                    x = cfg['x_col']
                                    ys = cfg['y_cols']
                                    r0, r1 = cfg['row_range']
                                    ctype = cfg['chart_type']
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
                                        chart.set_title({'name': f"{safe_name} - {ctype} (custom)"})
                                        chart.set_x_axis({'name': x})
                                        worksheet.insert_chart('G2', chart)
                                    elif ctype == 'Pie' and ys:
                                        val_col = ys[0]
                                        try:
                                            val_idx = list(df.columns).index(val_col)
                                            cat_idx = list(df.columns).index(x)
                                            pie = workbook.add_chart({'type': 'pie'})
                                            pie.add_series({
                                                'name': val_col,
                                                'categories': [safe_name, 1 + r0, cat_idx, 1 + r1, cat_idx],
                                                'values': [safe_name, 1 + r0, val_idx, 1 + r1, val_idx],
                                            })
                                            pie.set_title({'name': f"{safe_name} - Pie (custom)"})
                                            worksheet.insert_chart('G2', pie)
                                        except Exception:
                                            pass
                                except Exception as e:
                                    st.warning(f"Failed to insert custom chart for {safe_name}: {e}")
                            else:
                                # Auto-chart heuristics (same as preview): prefer Month multi-line
                                num_cols = df.select_dtypes(include=['number']).columns.tolist()
                                cat_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
                                if 'Month' in df.columns and num_cols:
                                    try:
                                        chart = workbook.add_chart({'type': 'line'})
                                        for i, col in enumerate(df.columns.tolist()):
                                            if col in num_cols:
                                                ci = df.columns.get_loc(col)
                                                chart.add_series({
                                                    'name': [safe_name, 0, ci],
                                                    'categories': [safe_name, 1, df.columns.get_loc('Month'), len(df), df.columns.get_loc('Month')],
                                                    'values': [safe_name, 1, ci, len(df), ci],
                                                })
                                        chart.set_title({'name': f"{safe_name} - Auto Trend"})
                                        chart.set_x_axis({'name': 'Month'})
                                        worksheet.insert_chart('G2', chart)
                                    except Exception:
                                        pass
                                elif num_cols and cat_cols:
                                    try:
                                        agg = df.groupby(cat_cols[0])[num_cols[0]].sum().reset_index()
                                        help_name = (safe_name + '_agg')[:31]
                                        agg.to_excel(writer, sheet_name=help_name, index=False)
                                        help_ws = writer.sheets[help_name]
                                        chart = workbook.add_chart({'type': 'column'})
                                        chart.add_series({
                                            'name': [help_name, 0, 1],
                                            'categories': [help_name, 1, 0, len(agg), 0],
                                            'values': [help_name, 1, 1, len(agg), 1],
                                        })
                                        chart.set_title({'name': f"{safe_name} - Auto {num_cols[0]} by {cat_cols[0]}"})
                                        worksheet.insert_chart('G2', chart)
                                    except Exception:
                                        pass
                                elif num_cols:
                                    try:
                                        chart = workbook.add_chart({'type': 'line'})
                                        ci = df.columns.get_loc(num_cols[0])
                                        chart.add_series({
                                            'name': [safe_name, 0, ci],
                                            'categories': [safe_name, 1, 0, len(df), 0],
                                            'values': [safe_name, 1, ci, len(df), ci],
                                        })
                                        chart.set_title({'name': f"{safe_name} - Auto {num_cols[0]}"})
                                        worksheet.insert_chart('G2', chart)
                                    except Exception:
                                        pass

                buffer.seek(0)
                out_path = REPORT_DIR / out_name
                with open(out_path, 'wb') as f:
                    f.write(buffer.getvalue())
                st.success(f"Report generated: {out_name} (stored in ./reports)")
                st.download_button('Download report', data=buffer, file_name=out_name, mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
                st.dataframe(list_saved_files(REPORT_DIR), use_container_width=True)
        except Exception as e:
            st.error(f"Report generation failed: {e}")

# ----- Past Reports tab -----
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
