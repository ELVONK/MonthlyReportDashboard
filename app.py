# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import re, time, uuid, json, shutil
import xlsxwriter
import plotly.express as px
import plotly.io as pio

st.set_page_config(page_title="Monthly Report Dashboard", layout="wide")

# -- simple theme colors (yellow/gold/green)
_THEME_CSS = """
:root{
  --brand-yellow: #FFF8E1;
  --brand-gold: #D4AF37;
  --brand-green: #2E7D32;
  --brand-dark: #263238;
}
.reportview-container, .main, .block-container { background-color: var(--brand-yellow) !important; }
.stButton>button, .stDownloadButton>button { background: linear-gradient(180deg, var(--brand-gold), #b58f2c) !important; color: white !important; border-radius: 10px !important; }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, #fffbe6, #f7f3e6) !important; }
.stAlert-success { background-color: rgba(46,125,50,0.08) !important; border-left: 4px solid var(--brand-green) !important; }
"""
st.markdown(f"<style>{_THEME_CSS}</style>", unsafe_allow_html=True)
st.markdown("<h1 style='color:#263238;margin-bottom:0.2rem'>Monthly Report Dashboard</h1>", unsafe_allow_html=True)
st.caption("Upload Excel/CSV files — preview, quick analysis, and auto-generated reports with embedded charts.")

# Paths
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
REPORT_DIR = BASE_DIR / "reports"
ARCHIVE_DIR = REPORT_DIR / "archive"
for d in (UPLOAD_DIR, REPORT_DIR, ARCHIVE_DIR):
    d.mkdir(parents=True, exist_ok=True)

PALETTE = ["#D4AF37", "#2E7D32", "#263238", "#F6E27A", "#9CCC65"]

# Helpers
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

# Excel engine chooser (requires openpyxl/xlrd)
try:
    import openpyxl  # noqa
    HAS_OPENPYXL = True
except Exception:
    HAS_OPENPYXL = False

try:
    import xlrd  # noqa
    HAS_XLRD = True
except Exception:
    HAS_XLRD = False

def _choose_excel_engine_from_filename(fname: str) -> dict:
    fname_l = fname.lower()
    if fname_l.endswith(".xlsx"):
        if not HAS_OPENPYXL:
            raise ImportError("openpyxl not installed")
        return {"engine": "openpyxl"}
    if fname_l.endswith(".xls"):
        if not HAS_XLRD:
            raise ImportError("xlrd not installed")
        return {"engine": "xlrd"}
    return {}

# Sidebar: upload & saved files selection
with st.sidebar:
    st.header("Files & Actions")
    uploaded_files = st.file_uploader("Upload Excel/CSV files", type=["xlsx", "xls", "csv"],
                                      accept_multiple_files=True)
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

    chosen_for_report = st.multiselect("Files to include in report", options=saved_options, default=None, key="chosen_for_report")
    st.markdown("---")
    if st.button("Open Reports Folder (in app) "):
        st.info(f"Reports are stored under: {REPORT_DIR}")

# Main UI tabs: Preview, Generate, Past
tabs = st.tabs(["Preview", "Generate Report", "Past Reports"])

# Quick analysis & auto-chart heuristic (used both in preview & report)
def quick_analysis_and_chart(df: pd.DataFrame, title: str):
    st.markdown(f"### {title}")
    if df.empty:
        st.info("Empty sheet")
        return

    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()

    st.write("Data sample:")
    st.dataframe(df.head(8), use_container_width=True)

    if num_cols:
        st.write("Summary stats for numeric columns:")
        st.dataframe(df[num_cols].describe().transpose(), use_container_width=True)
    else:
        st.info("No numeric columns found for detailed stats.")

    fig = None
    # Heuristics
    if "Month" in df.columns and num_cols:
        try:
            plot_df = df[["Month"] + num_cols].copy()
            for c in num_cols:
                plot_df[c] = pd.to_numeric(plot_df[c], errors="coerce")
            fig = px.line(plot_df, x="Month", y=num_cols, title=f"Auto: {title} — Multi-series over Month", markers=True, color_discrete_sequence=PALETTE)
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
            plot_df = df[num_cols].apply(pd.to_numeric, errors="coerce").reset_index()
            fig = px.line(plot_df, x="index", y=num_cols[0], title=f"Auto: {title} — {num_cols[0]}", color_discrete_sequence=PALETTE)
        except Exception:
            fig = None

    if fig is not None:
        st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True})
        html = pio.to_html(fig, full_html=True)
        st.download_button(label="Download auto-chart as HTML", data=html, file_name=f"{re.sub('[^0-9a-zA-Z]+','_', title)}_auto_chart.html", mime="text/html")
        # PNG optional (kaleido)
        try:
            img = fig.to_image(format="png")
            st.download_button(label="Download auto-chart as PNG", data=img, file_name=f"{re.sub('[^0-9a-zA-Z]+','_', title)}_auto_chart.png", mime="image/png")
        except Exception:
            st.caption("PNG export unavailable (install 'kaleido' to enable PNG export).")
    else:
        st.info("No suitable auto-chart could be generated for this sheet.")

# Preview tab
with tabs[0]:
    st.header("Preview files & Auto Analysis")
    st.markdown("Preview session uploads or saved files — the app will also run a quick analysis and draw auto-charts.")

    # session uploads first
    if uploaded_files:
        for uf in uploaded_files:
            with st.expander(f"(session) {uf.name}"):
                try:
                    fb = uf.getvalue()
                    if uf.name.lower().endswith(".csv"):
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

    st.markdown("---")
    st.subheader("Saved uploads analysis")
    saved_df = list_saved_files(UPLOAD_DIR)
    if not saved_df.empty:
        for fname in saved_df.sort_values("modified", ascending=False)["file"].tolist():
            with st.expander(fname):
                try:
                    fb = read_file_bytes_from_disk(UPLOAD_DIR / fname)
                    if fname.lower().endswith(".csv"):
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
        st.info("No saved uploads found yet.")

# Generate Report tab
with tabs[1]:
    st.header("Generate Report — Auto charts from uploaded files")
    st.markdown("Choose files in the sidebar and click Generate. The app will auto-generate charts per sheet and embed them in the Excel report.")
    base_name = st.text_input("Base report filename", value="Department_Report_with_Charts")
    generate = st.button("Generate Report (auto)")

    if generate:
        chosen = st.session_state.get("chosen_for_report", [])
        if not chosen:
            st.error("No files selected for the report. Pick files in the sidebar.")
        else:
            try:
                data_for_report = {}
                for fname in chosen:
                    p = UPLOAD_DIR / fname
                    if not p.exists():
                        st.warning(f"Skipping missing {fname}")
                        continue
                    fb = read_file_bytes_from_disk(p)
                    if fname.lower().endswith(".csv"):
                        data_for_report[fname] = {"(csv)": pd.read_csv(BytesIO(fb))}
                    else:
                        engine_kw = _choose_excel_engine_from_filename(fname)
                        xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                        data_for_report[fname] = {}
                        for s in xls.sheet_names:
                            data_for_report[fname][s] = pd.read_excel(BytesIO(fb), sheet_name=s, **engine_kw)

                if not data_for_report:
                    st.error("No valid data to write.")
                else:
                    ts = time.strftime("%Y%m%d-%H%M%S")
                    out_name = f"{base_name}_{ts}.xlsx"
                    buffer = BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        workbook = writer.book
                        for fname, sheets in data_for_report.items():
                            for sname, df in sheets.items():
                                safe_name = re.sub(r"[:\\/?*\[\]]", "_", f"{Path(fname).stem}_{sname}")[:31]
                                # write data
                                try:
                                    df.to_excel(writer, sheet_name=safe_name, index=False)
                                except Exception:
                                    df.head(1000).to_excel(writer, sheet_name=safe_name, index=False)
                                worksheet = writer.sheets[safe_name]

                                # auto-chart heuristics (same as preview)
                                num_cols = df.select_dtypes(include=["number"]).columns.tolist()
                                cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()

                                if "Month" in df.columns and num_cols:
                                    try:
                                        chart = workbook.add_chart({"type": "line"})
                                        for col in num_cols:
                                            ci = list(df.columns).index(col)
                                            chart.add_series({
                                                "name": [safe_name, 0, ci],
                                                "categories": [safe_name, 1, list(df.columns).index("Month"), len(df), list(df.columns).index("Month")],
                                                "values": [safe_name, 1, ci, len(df), ci],
                                            })
                                        chart.set_title({"name": f"{safe_name} - Auto Trend"})
                                        chart.set_x_axis({"name": "Month"})
                                        worksheet.insert_chart("G2", chart)
                                    except Exception:
                                        pass
                                elif num_cols and cat_cols:
                                    try:
                                        agg = df.groupby(cat_cols[0])[num_cols[0]].sum().reset_index()
                                        help_name = (safe_name + "_agg")[:31]
                                        agg.to_excel(writer, sheet_name=help_name, index=False)
                                        # ensure worksheet object exists
                                        help_ws = writer.sheets.get(help_name)
                                        if help_ws is None:
                                            help_ws = writer.book.add_worksheet(help_name)
                                            writer.sheets[help_name] = help_ws
                                        chart = workbook.add_chart({"type": "column"})
                                        chart.add_series({
                                            "name": [help_name, 0, 1],
                                            "categories": [help_name, 1, 0, len(agg), 0],
                                            "values": [help_name, 1, 1, len(agg), 1],
                                        })
                                        chart.set_title({"name": f"{safe_name} - Auto {num_cols[0]} by {cat_cols[0]}"})
                                        worksheet.insert_chart("G2", chart)
                                    except Exception:
                                        pass
                                elif num_cols:
                                    try:
                                        ci = list(df.columns).index(num_cols[0])
                                        chart = workbook.add_chart({"type": "line"})
                                        chart.add_series({
                                            "name": [safe_name, 0, ci],
                                            "categories": [safe_name, 1, 0, len(df), 0],
                                            "values": [safe_name, 1, ci, len(df), ci],
                                        })
                                        chart.set_title({"name": f"{safe_name} - Auto {num_cols[0]}"})
                                        worksheet.insert_chart("G2", chart)
                                    except Exception:
                                        pass

                    buffer.seek(0)
                    out_path = REPORT_DIR / out_name
                    with open(out_path, "wb") as f:
                        f.write(buffer.getvalue())

                    st.success(f"Report generated: {out_name} (stored in ./reports)")
                    st.download_button(label="⬇️ Download the Excel report", data=buffer, file_name=out_name,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    st.dataframe(list_saved_files(REPORT_DIR), use_container_width=True)
            except Exception as e:
                st.error(f"Failed to generate report: {e}")

# Past reports tab
with tabs[2]:
    st.header("Past reports")
    rep_df = list_saved_files(REPORT_DIR)
    if rep_df.empty:
        st.info("No reports yet.")
    else:
        st.dataframe(rep_df.sort_values("modified", ascending=False), use_container_width=True)
        pick = st.selectbox("Download report", options=rep_df.sort_values("modified", ascending=False)["file"].tolist())
        if pick:
            p = REPORT_DIR / pick
            if p.exists():
                with open(p, "rb") as f:
                    st.download_button(f"Download {pick}", data=f.read(), file_name=pick,
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.markdown("---")
    st.subheader("Manage reports")
    manage_choices = rep_df.sort_values("modified", ascending=False)["file"].tolist() if not rep_df.empty else []
    to_manage = st.multiselect("Select report(s) to archive or delete", options=manage_choices)
    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Archive selected"):
            if not to_manage:
                st.warning("No reports selected")
            else:
                for fn in to_manage:
                    src = REPORT_DIR / fn
                    dst = ARCHIVE_DIR / fn
                    try:
                        shutil.move(str(src), str(dst))
                    except Exception as e:
                        st.error(f"Failed to archive {fn}: {e}")
                st.success("Selected reports archived")
                st.experimental_rerun()
    with col_b:
        if st.button("Delete selected"):
            if not to_manage:
                st.warning("No reports selected")
            else:
                for fn in to_manage:
                    try:
                        (REPORT_DIR / fn).unlink()
                    except Exception as e:
                        st.error(f"Failed to delete {fn}: {e}")
                st.success("Selected reports deleted")
                st.experimental_rerun()
