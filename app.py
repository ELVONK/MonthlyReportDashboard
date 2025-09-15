# ---------------------
# Tabs (ensure they exist here)
# ---------------------
preview_tab, generate_tab, past_tab = st.tabs(["Preview", "Generate Report", "Past Reports"])

# ---------------------
# Generate Report tab (axis chooser + saved configs + generate)
# ---------------------
with generate_tab:
    st.header("Generate Report — choose axes for charts")
    st.markdown("Pick a sheet below to choose X and Y axes. Save axis configs and then Generate Report (auto charts used where no config exists).")

    base_name = st.text_input("Base report filename", value="Department_Report_with_Charts")
    generate = st.button("Generate Report (use saved axis configs)")

    # Ensure axis_config storage exists
    if "axis_config" not in st.session_state:
        st.session_state["axis_config"] = {}  # key -> dict {x_col, y_cols}

    # Read chosen files from sidebar session_state (sidebar multiselect must use same key)
    chosen_files = st.session_state.get("chosen_for_report", []) or []

    # Build available_sheets for chosen files
    def gather_available_sheets(selected_files: list) -> dict:
        out = {}
        for fname in (selected_files or []):
            p = UPLOAD_DIR / fname
            if not p.exists():
                continue
            try:
                fb = read_file_bytes_from_disk(p)
                if fname.lower().endswith(".csv"):
                    out[fname] = {"(csv)": pd.read_csv(BytesIO(fb))}
                else:
                    engine_kw = _choose_excel_engine_from_filename(fname)
                    xls = pd.ExcelFile(BytesIO(fb), **engine_kw)
                    out[fname] = {}
                    for s in xls.sheet_names:
                        out[fname][s] = pd.read_excel(BytesIO(fb), sheet_name=s, **engine_kw)
            except Exception:
                continue
        return out

    available_sheets = gather_available_sheets(chosen_files)
    sheet_choices = []
    sheet_map = {}
    for fname, sheets in available_sheets.items():
        for sname in sheets.keys():
            key = f"{fname}::{sname}"
            sheet_choices.append(key)
            sheet_map[key] = (fname, sname)

    if not sheet_choices:
        st.info("No sheets available. Select files in the sidebar and save uploaded files first.")
    else:
        chosen_sheet_key = st.selectbox("Pick a sheet to configure axes", options=[""] + sheet_choices, key="axis_pick")
        if chosen_sheet_key:
            fname, sname = sheet_map[chosen_sheet_key]
            df = available_sheets[fname][sname]
            st.markdown(f"**Preview: {fname} / {sname} (first 10 rows)**")
            st.dataframe(df.head(10), use_container_width=True)

            cols = list(df.columns)
            safe_key = re.sub(r"[^0-9a-zA-Z]+", "_", chosen_sheet_key)
            x_col = st.selectbox("Choose X-axis column (single)", options=[""] + cols, key=f"x_{safe_key}")
            y_cols = st.multiselect("Choose Y-axis column(s) (one or more)", options=cols, default=None, key=f"y_{safe_key}")

            # show existing
            existing = st.session_state["axis_config"].get(chosen_sheet_key)
            if existing:
                st.caption("Saved config for this sheet:")
                st.write(existing)

            c1, c2 = st.columns([1, 1])
            with c1:
                if st.button("Save axis config for this sheet", key=f"save_axis_{safe_key}"):
                    if not x_col or not y_cols:
                        st.warning("Select both X and at least one Y column before saving.")
                    else:
                        st.session_state["axis_config"][chosen_sheet_key] = {"x_col": x_col, "y_cols": y_cols}
                        st.success("Axis configuration saved.")
            with c2:
                if st.button("Remove saved config for this sheet", key=f"remove_axis_{safe_key}"):
                    if chosen_sheet_key in st.session_state["axis_config"]:
                        del st.session_state["axis_config"][chosen_sheet_key]
                        st.success("Removed saved configuration.")

            # Immediate preview if chosen
            if x_col and y_cols:
                try:
                    preview_df = df[[x_col] + y_cols].copy()
                    for yc in y_cols:
                        preview_df[yc] = pd.to_numeric(preview_df[yc], errors="coerce")
                    fig = px.line(preview_df, x=x_col, y=y_cols, markers=True,
                                  title=f"Preview: {Path(fname).stem}/{sname} — X: {x_col} Y: {', '.join(y_cols)}",
                                  color_discrete_sequence=PALETTE)
                    st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": True, "scrollZoom": True})
                    st.download_button("Download preview chart HTML", data=pio.to_html(fig, full_html=True),
                                       file_name=f"{Path(fname).stem}_{sname}_preview_chart.html", mime="text/html")
                except Exception as e:
                    st.warning(f"Could not render preview chart: {e}")

    # GENERATE: use stored axis_config if present; otherwise auto heuristics
    if generate:
        if not chosen_files:
            st.error("No files selected for the report. Pick files in the sidebar.")
        else:
            try:
                data_for_report = {}
                for fname in chosen_files:
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

                ts = time.strftime("%Y%m%d-%H%M%S")
                out_name = f"{base_name}_{ts}.xlsx"
                buffer = BytesIO()
                with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                    workbook = writer.book
                    for fname, sheets in data_for_report.items():
                        for sname, df in sheets.items():
                            safe_name = re.sub(r"[:\\/?*\[\]]", "_", f"{Path(fname).stem}_{sname}")[:31]
                            try:
                                df.to_excel(writer, sheet_name=safe_name, index=False)
                            except Exception:
                                df.head(1000).to_excel(writer, sheet_name=safe_name, index=False)
                            worksheet = writer.sheets[safe_name]

                            sheet_key = f"{fname}::{sname}"
                            cfg = st.session_state.get("axis_config", {}).get(sheet_key)

                            if cfg:
                                # embed chart using configured axes
                                x = cfg["x_col"]
                                ys = cfg["y_cols"]
                                try:
                                    chart = workbook.add_chart({"type": "line"})
                                    rcount = max(1, len(df))
                                    for col in ys:
                                        try:
                                            col_idx = list(df.columns).index(col)
                                            chart.add_series({
                                                "name": [safe_name, 0, col_idx],
                                                "categories": [safe_name, 1, list(df.columns).index(x), 1 + rcount - 1, list(df.columns).index(x)],
                                                "values": [safe_name, 1, col_idx, 1 + rcount - 1, col_idx],
                                            })
                                        except Exception:
                                            continue
                                    chart.set_title({"name": f"{safe_name} - Custom axes"})
                                    chart.set_x_axis({"name": x})
                                    worksheet.insert_chart("G2", chart)
                                except Exception as e:
                                    st.warning(f"Failed to embed custom chart for {safe_name}: {e}")
                            else:
                                # fallback to auto heuristics (same as preview logic)
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
