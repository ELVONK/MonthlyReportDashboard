# app.py
import streamlit as st
import pandas as pd
from io import BytesIO
from pathlib import Path
import re, time, uuid
import xlsxwriter  # required for Excel writer engine

st.set_page_config(page_title="Monthly Report Dashboard", layout="wide")
st.title("Monthly Report Dashboard — Upload & Save Files")

# ---------- Utilities ----------
UPLOAD_DIR = Path("user_uploads")
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

def safe_unique_name(original_name: str) -> str:
    p = Path(original_name)
    stem = re.sub(r"[^A-Za-z0-9._-]+", "_", p.stem)[:80] or "file"
    suffix = p.suffix.lower()
    ts = time.strftime("%Y%m%d-%H%M%S")
    uid = uuid.uuid4().hex[:6]
    return f"{stem}_{ts}_{uid}{suffix}"

def save_uploaded_file(up_file) -> Path:
    """Save a Streamlit UploadedFile to disk with a safe unique name."""
    dest = UPLOAD_DIR / safe_unique_name(up_file.name)
    with open(dest, "wb") as f:
        f.write(up_file.getbuffer())
    return dest

def list_saved():
    rows = []
    for p in sorted(UPLOAD_DIR.glob("*")):
        if p.is_file():
            rows.append({"file": p.name, "size_kb": round(p.stat().st_size/1024, 1)})
    return pd.DataFrame(rows)

# ---------- Sidebar controls ----------
with st.sidebar:
    st.header("Upload")
    uploaded_files = st.file_uploader(
        "Upload Excel/CSV files",
        type=["xlsx", "xls", "csv"],
        accept_multiple_files=True,
        help="All uploads will be saved to server under user_uploads/ with unique names.",
    )
    save_btn = st.button("Save all uploaded files to disk", use_container_width=True)

# ---------- Always render something helpful ----------
st.subheader("Saved Files")
df_saved = list_saved()
st.dataframe(df_saved, use_container_width=True, height=220)

# ---------- Handle uploads & saving ----------
if save_btn:
    if not uploaded_files:
        st.warning("No files selected. Please upload at least one file.")
    else:
        saved_paths = []
        for uf in uploaded_files:
            try:
                path = save_uploaded_file(uf)
                saved_paths.append(path)
            except Exception as e:
                st.error(f"Failed to save **{uf.name}**: {e}")
        if saved_paths:
            st.success(f"Saved {len(saved_paths)} file(s) to **{UPLOAD_DIR}/**.")
            st.dataframe(
                pd.DataFrame({"saved_path": [str(p) for p in saved_paths]}),
                use_container_width=True,
            )
            # refresh table
            st.experimental_rerun()

# ---------- Your report generation (kept, plus saved to disk) ----------
st.subheader("Generate Department Report with Charts")
st.caption("This uses your in-app data dict, writes charts, saves to disk, and provides a download.")

# Your original in-memory data:
data = {
    "Supply Chain": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Purchases Made": [25, 30, 28],
        "Procurement Plan Monitoring": [80, 85, 90],
        "Special Group Contracts": [5, 6, 4],
        "Suppliers Registered": [10, 15, 12]
    }),
    "Human Resources": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Staff Training": [2, 3, 4],
        "Staff Welfare Activities": [1, 2, 2],
        "Complaints Received": [3, 4, 2],
        "Complaints Resolved": [2, 4, 2],
        "Students on Industrial Training": [5, 6, 7]
    }),
    "Road Asset and Corridor Management": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Road Works Progress (%)": [60, 75, 85],
        "Inspections Done": [8, 10, 12],
        "Achievements Reported": [5, 6, 7]
    }),
    "Transport": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Service and Maintenance": [10, 12, 9],
        "Fuel Consumption (Litres)": [500, 600, 550]
    }),
    "Survey": pd.DataFrame({
        "Month": ["Jan", "Feb", "Mar"],
        "Surveys Completed": [10, 12, 14],
        "Pending Reports": [2, 1, 0]
    }),
    "Finance and Accounts": pd.DataFrame({
        "Contracts Paid": ["Contract A", "Contract B", "Contract C"],
        "Amount Paid": [10000, 15000, 12000],
        "Per Diem Paid": [3000, 2500, 2700],
        "Budget Consumption": [18000, 20000, 19000]
    })
}

sheet_names_fixed = {
    "Supply Chain": "Supply Chain",
    "Human Resources": "HR",
    "Road Asset and Corridor Management": "Road Assets",
    "Transport": "Transport",
    "Survey": "Survey",
    "Finance and Accounts": "Finance"
}

col1, col2 = st.columns([1, 1])
with col1:
    out_filename = st.text_input(
        "Output Excel filename",
        value="Department_Report_with_Charts.xlsx",
        help="Will be saved under user_uploads/ and also provided as a download.",
    )
with col2:
    gen_btn = st.button("Generate Report", type="primary", use_container_width=True)

if gen_btn:
    try:
        # Write once to a BytesIO (for download) and also to disk
        buffer = BytesIO()
        # Write to BytesIO
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            workbook = writer.book
            for dept, df in data.items():
                sheet_name = sheet_names_fixed[dept]
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                if "Month" in df.columns:
                    chart = workbook.add_chart({"type": "line"})
                    for i, col in enumerate(df.columns[1:], start=1):
                        chart.add_series({
                            "name":       [sheet_name, 0, i],
                            "categories": [sheet_name, 1, 0, len(df), 0],
                            "values":     [sheet_name, 1, i, len(df), i],
                        })
                    chart.set_title({"name": f"{sheet_name} - Multi-Month Trend"})
                    chart.set_x_axis({"name": "Month"})
                    chart.set_y_axis({"name": "Value"})
                    worksheet.insert_chart("G2", chart)

        buffer.seek(0)

        # Save the same bytes to disk under user_uploads/
        out_path = UPLOAD_DIR / safe_unique_name(out_filename)
        with open(out_path, "wb") as f:
            f.write(buffer.getvalue())

        st.success(f"Report generated and saved as **{out_path.name}** in **{UPLOAD_DIR}/**.")

        st.download_button(
            label="⬇️ Download the Excel report",
            data=buffer,
            file_name=out_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Refresh saved list table
        st.dataframe(list_saved(), use_container_width=True, height=220)

    except Exception as e:
        st.error(f"Failed to generate the report: {e}")

# ---------- (Optional) S3 persistence stub ----------
with st.expander("Optional: Save uploads to Amazon S3 (advanced)"):
    st.markdown(
        """
        If you need durable storage, consider S3. Example snippet:

        ```python
        import boto3, os
        s3 = boto3.client("s3",
                          aws_access_key_id=st.secrets["AWS_ACCESS_KEY_ID"],
                          aws_secret_access_key=st.secrets["AWS_SECRET_ACCESS_KEY"],
                          region_name=st.secrets["AWS_REGION"])
        bucket = st.secrets["S3_BUCKET"]

        for uf in uploaded_files:
            key = f"uploads/{safe_unique_name(uf.name)}"
            s3.put_object(Bucket=bucket, Key=key, Body=uf.getbuffer())
        ```
        """
    )
