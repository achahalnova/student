import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
import os

st.set_page_config(page_title="Student Excel Analyzer", page_icon="📊", layout="wide")

# Sidebar Navigation instead of Tabs (much more stable for older Streamlit versions)
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to", ["Results Analyzer (Excel)", "Exam Grader (.dat)"])

if page == "Results Analyzer (Excel)":
    st.title("Student Results Analyzer 📊")
    st.write(
        "Upload your weekly JEE or NEET Excel/CSV file. "
        "This app filters your selected branch and color-codes **subject & total ranks per branch** in accurate 20% tiers "
        "(Top 20% = Green, 20–40% = Light Green, 40–60% = Yellow, 60–80% = Orange, Bottom 20% = Red)."
    )

    exam_type = st.radio("Select Exam Type:", ["Engineering (JEE)", "Medical (NEET)"])
    uploaded_file = st.file_uploader("Upload Student Results Excel", type=["xlsx", "xls", "csv"])

    if uploaded_file is not None:
        st.info("Reading file...")

        try:
            if uploaded_file.name.lower().endswith(".csv"):
                try:
                    df_raw = pd.read_csv(uploaded_file, header=None)
                except UnicodeDecodeError:
                    uploaded_file.seek(0)
                    df_raw = pd.read_csv(uploaded_file, header=None, encoding="latin1")
            else:
                if exam_type == "Medical (NEET)":
                    try:
                        df_raw = pd.read_excel(uploaded_file, sheet_name="EST", header=None)
                    except Exception:
                        st.warning("Could not find 'EST' sheet. Falling back to the first sheet...")
                        uploaded_file.seek(0)
                        df_raw = pd.read_excel(uploaded_file, sheet_name=0, header=None)
                else:
                    df_raw = pd.read_excel(uploaded_file, header=None)

            if len(df_raw) < 3:
                st.error("Error: Excel file does not contain enough rows to analyze.")
                st.stop()

            # Dynamic header row detection
            header_row_idx = 1
            for idx in range(min(15, len(df_raw))):
                row_vals = [str(x).upper().strip() for x in df_raw.iloc[idx].tolist()]
                if "NAME" in row_vals or "STUDENT NAME" in row_vals or "BRANCH" in row_vals:
                    header_row_idx = idx
                    break

            data_rows = df_raw.iloc[header_row_idx + 1:].copy()
            header_row = [str(x).strip() for x in df_raw.iloc[header_row_idx].tolist()]
            header_upper = [h.upper().strip() for h in header_row]

            def get_first_idx(col_name_or_list, start_idx=0, end_idx=None, required=True):
                col_names = [
                    str(c).upper().strip()
                    for c in (col_name_or_list if isinstance(col_name_or_list, list) else [col_name_or_list])
                ]
                stop_at = len(header_upper) if end_idx is None else min(end_idx, len(header_upper))
                for i in range(start_idx, stop_at):
                    if header_upper[i] in col_names:
                        return i
                if required:
                    available_cols = ", ".join(
                        [h for h in header_row if h and h.lower() != "nan"]
                    )
                    st.error(
                        f"Error: Could not find column '{col_name_or_list}' in the header. "
                        f"Available columns: {available_cols}"
                    )
                    st.stop()
                return None

            branch_idx = get_first_idx(["BRANCH", "CAMPUS", "CENTRE", "CENTER"], required=False)
            name_idx = get_first_idx(["NAME", "STUDENT NAME", "STUDENT_NAME"])

            pm_idx = get_first_idx("PM", 0)
            cm_idx = get_first_idx("CM", 0)
            tm_idx = get_first_idx(["TM", "TOTAL", "TOTAL MARKS"], 0)

            if exam_type == "Medical (NEET)":
                bm_idx = get_first_idx("BM", 0)
                zm_idx = get_first_idx("ZM", 0)
                bio_m_idx = get_first_idx(["BIO-M", "BIOM", "BIOLOGY", "BIO"], 0, required=False)

                subj_starts = sorted([i for i in [pm_idx, cm_idx, bm_idx, zm_idx, bio_m_idx, tm_idx] if i is not None])
                def next_boundary(idx):
                    for s in subj_starts:
                        if s > idx:
                            return s
                    return len(header_upper)

                pr_idx = get_first_idx("PR", pm_idx, next_boundary(pm_idx), required=False)
                pw_idx = get_first_idx("W", pm_idx, next_boundary(pm_idx), required=False)

                cr_idx = get_first_idx("CR", cm_idx, next_boundary(cm_idx), required=False)
                cw_idx = get_first_idx("W", cm_idx, next_boundary(cm_idx), required=False)

                br_idx = get_first_idx("BR", bm_idx, next_boundary(bm_idx), required=False)
                bw_idx = get_first_idx("W", bm_idx, next_boundary(bm_idx), required=False)

                zr_idx = get_first_idx("ZR", zm_idx, next_boundary(zm_idx), required=False)
                zw_idx = get_first_idx("W", zm_idx, next_boundary(zm_idx), required=False)

                tr_idx = get_first_idx("TR", tm_idx, len(header_upper), required=False)
            else:
                mm_idx = get_first_idx("MM", 0)
                subj_starts = sorted([i for i in [mm_idx, pm_idx, cm_idx, tm_idx] if i is not None])
                def next_boundary(idx):
                    for s in subj_starts:
                        if s > idx:
                            return s
                    return len(header_upper)

                mr_idx = get_first_idx("MR", mm_idx, next_boundary(mm_idx), required=False)
                mw_idx = get_first_idx("W", mm_idx, next_boundary(mm_idx), required=False)

                pr_idx = get_first_idx("PR", pm_idx, next_boundary(pm_idx), required=False)
                pw_idx = get_first_idx("W", pm_idx, next_boundary(pm_idx), required=False)

                cr_idx = get_first_idx("CR", cm_idx, next_boundary(cm_idx), required=False)
                cw_idx = get_first_idx("W", cm_idx, next_boundary(cm_idx), required=False)

                tr_idx = get_first_idx("TR", tm_idx, len(header_upper), required=False)

            # Remove footer summary blocks (e.g., MARKS SUMMARY / ZONE SUMMARY) and empty/header rows
            summary_markers = ("MARKS SUMMARY", "ZONE SUMMARY")
            first_summary_pos = None
            for pos_idx, (_, row_series) in enumerate(data_rows.iterrows()):
                row_text = " ".join(str(v).upper().strip() for v in row_series.tolist() if pd.notna(v))
                if any(marker in row_text for marker in summary_markers):
                    first_summary_pos = pos_idx
                    break
            if first_summary_pos is not None:
                data_rows = data_rows.iloc[:first_summary_pos].copy()

            data_rows = data_rows.copy()
            name_clean = data_rows.iloc[:, name_idx].astype(str).str.strip()
            invalid_names = {"", "NAN", "NONE", "NAME", "STUDENT NAME", "MARKS SUMMARY", "ZONE SUMMARY", "STRENGTH", "TOTAL"}
            valid_name_mask = data_rows.iloc[:, name_idx].notna() & (~name_clean.str.upper().isin(invalid_names))
            data_rows = data_rows[valid_name_mask].copy()

            if branch_idx is not None:
                data_rows.iloc[:, branch_idx] = data_rows.iloc[:, branch_idx].apply(
                    lambda v: str(v).strip() if pd.notna(v) and str(v).strip().lower() != "nan" else None
                )

            include_branch_col = False
            if branch_idx is not None:
                all_branches = sorted([b for b in data_rows.iloc[:, branch_idx].dropna().unique().tolist() if b])
                mh_mum_branches = [b for b in all_branches if b.startswith("MH-MUM")]
                if "MH-MUM-NSP-NSPIRA-CC" in mh_mum_branches:
                    mh_mum_branches.remove("MH-MUM-NSP-NSPIRA-CC")
                    mh_mum_branches.insert(0, "MH-MUM-NSP-NSPIRA-CC")

                show_all_branches_list = st.checkbox("Show all branches in dropdown (not just MH-MUM)", value=False)
                if show_all_branches_list or not mh_mum_branches:
                    base_branches = all_branches
                else:
                    base_branches = mh_mum_branches

                if not base_branches:
                    st.warning("No branches were found in the uploaded file. Showing all students.")
                    df_filtered = data_rows.copy()
                else:
                    special_opts = []
                    if mh_mum_branches and not show_all_branches_list:
                        special_opts.append("All MH-MUM Branches (Color-coded per branch)")
                    special_opts.append("All Branches (Color-coded per branch)")

                    branch_options = base_branches + special_opts
                    target_branch = st.selectbox("Select branch to filter by:", options=branch_options)

                    if target_branch == "All MH-MUM Branches (Color-coded per branch)":
                        st.info("Showing all MH-MUM branches (color-coded per branch)...")
                        df_filtered = data_rows[data_rows.iloc[:, branch_idx].isin(mh_mum_branches)].copy()
                        include_branch_col = True
                    elif target_branch == "All Branches (Color-coded per branch)":
                        st.info("Showing all branches (color-coded per branch)...")
                        df_filtered = data_rows.copy()
                        include_branch_col = True
                    else:
                        st.info(f"Filtering for branch {target_branch}...")
                        df_filtered = data_rows[data_rows.iloc[:, branch_idx] == target_branch].copy()

                    if df_filtered.empty:
                        st.warning(f"No students found for selection '{target_branch}'.")
                        st.stop()
            else:
                st.info("No BRANCH column found. Showing all students.")
                df_filtered = data_rows.copy()

            # Build output columns dynamically based on columns present in the file
            col_spec = [("NAME", name_idx, False)]
            if include_branch_col and branch_idx is not None:
                col_spec.append(("BRANCH", branch_idx, False))

            if exam_type == "Medical (NEET)":
                candidate_cols = [
                    ("PM", pm_idx, True),
                    ("PR", pr_idx, True),
                    ("Physics W", pw_idx, True),
                    ("CM", cm_idx, True),
                    ("CR", cr_idx, True),
                    ("Chemistry W", cw_idx, True),
                    ("BM", bm_idx, True),
                    ("BR", br_idx, True),
                    ("Botany W", bw_idx, True),
                    ("ZM", zm_idx, True),
                    ("ZR", zr_idx, True),
                    ("Zoology W", zw_idx, True),
                    ("BIO-M", bio_m_idx, True),
                    ("TM", tm_idx, True),
                    ("TR", tr_idx, True),
                ]
                rank_cols_to_color = ["PR", "CR", "BR", "ZR", "TR"]
            else:
                candidate_cols = [
                    ("MM", mm_idx, True),
                    ("MR", mr_idx, True),
                    ("Math W", mw_idx, True),
                    ("PM", pm_idx, True),
                    ("PR", pr_idx, True),
                    ("Physics W", pw_idx, True),
                    ("CM", cm_idx, True),
                    ("CR", cr_idx, True),
                    ("Chemistry W", cw_idx, True),
                    ("TM", tm_idx, True),
                    ("TR", tr_idx, True),
                ]
                rank_cols_to_color = ["MR", "PR", "CR", "TR"]

            for col_label, c_idx, is_num in candidate_cols:
                if c_idx is not None:
                    col_spec.append((col_label, c_idx, is_num))

            req_indices = [item[1] for item in col_spec]
            out_col_names = [item[0] for item in col_spec]
            cols_to_convert = [item[0] for item in col_spec if item[2]]

            # Track branch key per row even when BRANCH column isn't exported
            if branch_idx is not None:
                branch_keys = df_filtered.iloc[:, branch_idx].fillna("UNKNOWN").astype(str).reset_index(drop=True)
            else:
                branch_keys = pd.Series(["ALL"] * len(df_filtered))

            df_final = df_filtered.iloc[:, req_indices].copy()
            df_final.columns = out_col_names
            df_final = df_final.reset_index(drop=True)

            for col in cols_to_convert:
                df_final[col] = pd.to_numeric(df_final[col], errors="coerce")

            # Sort students by Total Rank (TR) best-to-worst (and by BRANCH when multiple branches are shown)
            df_final["_branch_key"] = branch_keys
            if "TR" in df_final.columns:
                sort_cols = ["_branch_key", "TR"] if include_branch_col else ["TR"]
                df_final = df_final.sort_values(by=sort_cols, ascending=True, na_position="last").reset_index(drop=True)
            elif "TM" in df_final.columns:
                if include_branch_col:
                    df_final = df_final.sort_values(by=["_branch_key", "TM"], ascending=[True, False], na_position="last").reset_index(drop=True)
                else:
                    df_final = df_final.sort_values(by="TM", ascending=False, na_position="last").reset_index(drop=True)

            branch_keys = df_final.pop("_branch_key").reset_index(drop=True)

            # Midpoint per-branch rank percentile tiers (lower rank = better):
            # Uses (average_rank - 0.5) / n so ties always share a color, multiples of 5 split evenly,
            # and small branches (n=1, 2, 3, ...) map cleanly to Green -> Red without skipping bottom tiers.
            def get_branch_rank_tiers(rank_series, group_keys):
                tiers = {0: [], 1: [], 2: [], 3: [], 4: []}
                temp = pd.DataFrame({"rank_val": rank_series, "branch": group_keys}).dropna(subset=["rank_val"])
                if temp.empty:
                    return tiers

                for _, grp in temp.groupby("branch", sort=False):
                    n = len(grp)
                    if n == 0:
                        continue
                    if n == 1:
                        tiers[0].append(grp.index[0])
                        continue
                    if n == 2:
                        min_val = grp["rank_val"].min()
                        max_val = grp["rank_val"].max()
                        if min_val == max_val:
                            for row_idx in grp.index:
                                tiers[0].append(row_idx)
                        else:
                            for row_idx, val in grp["rank_val"].items():
                                tiers[0 if val == min_val else 4].append(row_idx)
                        continue

                    in_branch_rank = grp["rank_val"].rank(ascending=True, method="average")
                    midpoint_pct = (in_branch_rank - 0.5) / float(n)
                    for row_idx, pct_val in midpoint_pct.items():
                        if pct_val <= 0.20:
                            tiers[0].append(row_idx)
                        elif pct_val <= 0.40:
                            tiers[1].append(row_idx)
                        elif pct_val <= 0.60:
                            tiers[2].append(row_idx)
                        elif pct_val <= 0.80:
                            tiers[3].append(row_idx)
                        else:
                            tiers[4].append(row_idx)
                return tiers

            st.info("Calculating accurate per-branch rank percentiles (20% tiers) and coloring...")

            output_buffer = io.BytesIO()
            df_final.to_excel(output_buffer, index=False)
            output_buffer.seek(0)

            wb = openpyxl.load_workbook(output_buffer)
            ws = wb.active

            ws.column_dimensions["A"].width = 28

            tier_fills = {
                0: PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid"),  # Top 20%: Green
                1: PatternFill(start_color="CCFF99", end_color="CCFF99", fill_type="solid"),  # 20-40%: Light Green
                2: PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid"),  # 40-60%: Yellow
                3: PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid"),  # 60-80%: Orange
                4: PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),  # Bottom 20%: Red
            }

            def apply_tier_colors(ws, col_name, tiers_dict):
                col_idx = None
                for cell in ws[1]:
                    if cell.value == col_name:
                        col_idx = cell.column
                        break
                if not col_idx:
                    return
                for tier_id, row_indices in tiers_dict.items():
                    fill = tier_fills[tier_id]
                    for idx in row_indices:
                        ws.cell(row=idx + 2, column=col_idx).fill = fill

            rank_tiers_by_col = {}
            for rank_col in rank_cols_to_color:
                if rank_col in df_final.columns:
                    col_tiers = get_branch_rank_tiers(df_final[rank_col], branch_keys)
                    rank_tiers_by_col[rank_col] = col_tiers
                    apply_tier_colors(ws, rank_col, col_tiers)

            # Add a Color Legend & Branch Rank Cutoffs sheet for transparent review
            legend_ws = wb.create_sheet(title="Color Legend")
            legend_ws.append(["Tier (Per Branch by Rank)", "Color", "Students (TR)" if "TR" in rank_tiers_by_col else "Method"])
            legend_rows = [
                ("Top 20% (0% - 20%)", 0, "Green"),
                ("Next 20% (20% - 40%)", 1, "Light Green"),
                ("Middle 20% (40% - 60%)", 2, "Yellow"),
                ("Next 20% (60% - 80%)", 3, "Orange"),
                ("Bottom 20% (80% - 100%)", 4, "Red"),
            ]
            for r_i, (label, t_id, color_name) in enumerate(legend_rows, start=2):
                legend_ws.cell(row=r_i, column=1, value=label)
                c = legend_ws.cell(row=r_i, column=2, value=color_name)
                c.fill = tier_fills[t_id]
                if "TR" in rank_tiers_by_col:
                    legend_ws.cell(row=r_i, column=3, value=len(rank_tiers_by_col["TR"][t_id]))

            if "TR" in rank_tiers_by_col:
                legend_ws.append([])
                legend_ws.append(["Branch", "Strength", "Green (Top 20%)", "Light Green (20-40%)", "Yellow (40-60%)", "Orange (60-80%)", "Red (Bottom 20%)"])
                row_tier_map = {}
                for t_id, idxs in rank_tiers_by_col["TR"].items():
                    for i in idxs:
                        row_tier_map[i] = t_id
                df_summary = pd.DataFrame({"branch": branch_keys, "TR": df_final["TR"]})
                df_summary["tier"] = df_summary.index.map(row_tier_map)
                for b_name, b_grp in df_summary.dropna(subset=["TR"]).groupby("branch", sort=False):
                    tier_ranges = []
                    for t_id in range(5):
                        vals = b_grp[b_grp["tier"] == t_id]["TR"]
                        if vals.empty:
                            tier_ranges.append("-")
                        else:
                            lo, hi = int(vals.min()), int(vals.max())
                            tier_ranges.append(f"{lo}" if lo == hi else f"{lo} - {hi}")
                    legend_ws.append([b_name, len(b_grp)] + tier_ranges)

            legend_ws.column_dimensions["A"].width = 30
            for col_let in ["B", "C", "D", "E", "F", "G"]:
                legend_ws.column_dimensions[col_let].width = 18

            final_buffer = io.BytesIO()
            wb.save(final_buffer)
            final_buffer.seek(0)

            st.success(f"Analysis Complete! Processed {len(df_final)} student(s).")
            st.caption(
                "**Per-Branch Color Key:** 🟩 Green = Top 20% | 🟢 Light Green = 20–40% | 🟨 Yellow = 40–60% | 🟧 Orange = 60–80% | 🟥 Red = Bottom 20%"
            )

            st.write("### Data Preview (First 10 Rows)")
            st.dataframe(df_final.head(10))

            base_name, _ = os.path.splitext(uploaded_file.name)
            new_filename = f"{base_name}_Analyzed.xlsx"
            st.download_button(
                label="⬇️ Download Formatted Excel File",
                data=final_buffer,
                file_name=new_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")

elif page == "Exam Grader (.dat)":
    st.title("Exam Grader (.dat) 📝")
    st.write("Upload student .dat files and an Answer Key to process scores.")
    
    batch_options = [
        "junior jee 1", "junior jee 2", "junior jee 3", "junior jee 4", "junior jee 5",
        "junior neet 1", "junior neet 2", "junior neet 3",
        "senior jee 1", "senior jee 2", "senior jee 3", "senior jee 4", "senior jee 5",
        "senior neet 1", "senior neet 2",
        "start neet"
    ]
    
    selected_batch = st.selectbox("Select Batch:", options=batch_options)
    
    st.write("### Upload Files")
    col1, col2 = st.columns(2)
    with col1:
        dat_file = st.file_uploader("Upload Student Data File (.dat/.iit or extensionless)", type=None)
    with col2:
        answer_key_file = st.file_uploader("Upload Answer Key", type=["docx", "doc", "xlsx", "xls"])
    
    if dat_file is not None and answer_key_file is not None:
        st.info("Files uploaded successfully.")
        
        def extract_answer_key_from_docx(docx_file):
            import docx
            import pandas as pd
            
            doc = docx.Document(docx_file)
            data = []
        
            for table_idx, table in enumerate(doc.tables):
                if table_idx == 0:
                    first_cell_text = table.rows[0].cells[0].text.strip()
                    if not first_cell_text.isdigit():
                        continue
        
                for i in range(0, len(table.rows), 2):
                    if i + 1 >= len(table.rows):
                        break
                    q_row = table.rows[i]
                    a_row = table.rows[i+1]
                    
                    for q_cell, a_cell in zip(q_row.cells, a_row.cells):
                        q_text = q_cell.text.strip()
                        a_text = a_cell.text.strip()
                        
                        if q_text.isdigit():
                            data.append({
                                "Question Number": int(q_text),
                                "Answer": a_text
                            })
        
            if not data:
                return None
        
            data.sort(key=lambda x: x["Question Number"])
            return pd.DataFrame(data)

        # Process the Answer Key File
        if answer_key_file.name.lower().endswith(".docx"):
            st.info("Extracting answer key from Word Document...")
            try:
                df_answer_key = extract_answer_key_from_docx(answer_key_file)
                
                if df_answer_key is not None:
                    st.success("Answer key extracted successfully!")
                    
                    st.write("### Answer Key Preview")
                    st.dataframe(df_answer_key.head(10))
                    
                    ak_buffer = io.BytesIO()
                    df_answer_key.to_excel(ak_buffer, index=False)
                    ak_buffer.seek(0)
                    
                    st.download_button(
                        label="⬇️ Download Converted Answer Key (Excel)",
                        data=ak_buffer,
                        file_name="Converted_Answer_Key.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    st.error("Could not find a valid answer key table in the Word document.")
            except Exception as e:
                st.error(f"Error processing Word document: {e}. Note: .doc files are not supported, please convert to .docx")
        
        elif answer_key_file.name.lower().endswith(".xlsx") or answer_key_file.name.lower().endswith(".xls"):
            try:
                df_answer_key = pd.read_excel(answer_key_file)
                st.success("Excel answer key loaded successfully!")
                st.write("### Answer Key Preview")
                st.dataframe(df_answer_key.head(10))
            except Exception as e:
                st.error(f"Error reading Excel file: {e}")
                
        # Process the DAT File against the Answer Key
        if df_answer_key is not None:
            if st.button("Grade DAT File"):
                if "jee" in selected_batch.lower():
                    if True:
                        st.info("Grading student DAT file against the Answer Key...")
                    
                    try:
                        # Convert df_ak answers to a dictionary: {q_num: answer}
                        # map A=1, B=2, C=3, D=4
                        ans_map = {'A':'1', 'B':'2', 'C':'3', 'D':'4'}
                        ak_dict = {}
                        for _, row in df_answer_key.iterrows():
                            q_num = int(row["Question Number"])
                            ans_raw = str(row["Answer"]).strip().upper()
                            ans_val = ans_map.get(ans_raw, ans_raw)
                            ak_dict[q_num] = ans_val
                            
                        results = []
                        
                        # Read DAT file
                        content = dat_file.read().decode('utf-8').splitlines()
                        for line in content:
                            line = line.strip()
                            if not line:
                                continue
                            
                            # Split by comma or space
                            parts = [p.strip() for p in line.replace(',', ' ').split() if p.strip()]
                            if len(parts) < 77:
                                continue # Malformed row
                                
                            student_id = parts[1]
                            responses = parts[2:77] # 75 responses
                            
                            # Sub-sections
                            phys_obj = responses[0:20]
                            phys_int = responses[20:25]
                            chem_obj = responses[25:45]
                            chem_int = responses[45:50]
                            math_obj = responses[50:70]
                            math_int = responses[70:75]
                            
                            def score_section(obj_responses, int_responses, start_q):
                                marks = 0
                                wrong = 0
                                unattempted = 0
                                
                                # Objective
                                for i, r in enumerate(obj_responses):
                                    q = start_q + i
                                    correct_ans = ak_dict.get(q, None)
                                    if r == '0':
                                        unattempted += 1
                                    elif correct_ans and r == str(correct_ans):
                                        marks += 4
                                    else:
                                        marks -= 1
                                        wrong += 1
                                        
                                # Integer
                                for i, r in enumerate(int_responses):
                                    q = start_q + 20 + i
                                    correct_ans = ak_dict.get(q, None)
                                    
                                    if "-10000" in r:
                                        unattempted += 1
                                        continue
                                        
                                    if correct_ans:
                                        try:
                                            r_float = float(r)
                                            c_float = float(correct_ans)
                                            if r_float == c_float:
                                                marks += 4
                                            else:
                                                marks -= 1
                                                wrong += 1
                                        except ValueError:
                                            if r == str(correct_ans):
                                                marks += 4
                                            else:
                                                marks -= 1
                                                wrong += 1
                                    else:
                                        unattempted += 1
                    
                                return marks, wrong, unattempted
                    
                            p_m, p_w, p_u = score_section(phys_obj, phys_int, 1)
                            c_m, c_w, c_u = score_section(chem_obj, chem_int, 26)
                            m_m, m_w, m_u = score_section(math_obj, math_int, 51)
                            
                            total = p_m + c_m + m_m
                            
                            results.append({
                                "Student Code": student_id,
                                "Physics Marks": p_m,
                                "Physics Wrong": p_w,
                                "Physics Unattempted": p_u,
                                "Chemistry Marks": c_m,
                                "Chemistry Wrong": c_w,
                                "Chemistry Unattempted": c_u,
                                "Math Marks": m_m,
                                "Math Wrong": m_w,
                                "Math Unattempted": m_u,
                                "Total Marks": total
                            })
                            
                        df_results = pd.DataFrame(results)
                        
                        # Sort by Total Marks descending
                        if not df_results.empty:
                            df_results = df_results.sort_values(by="Total Marks", ascending=False).reset_index(drop=True)
                        
                        st.success("DAT file graded successfully!")
                        st.write("### Graded Results Preview")
                        st.dataframe(df_results.head(10))
                        
                        import openpyxl
                        from openpyxl.styles import PatternFill

                        out_buffer = io.BytesIO()
                        df_results.to_excel(out_buffer, index=False)
                        out_buffer.seek(0)
                        
                        wb = openpyxl.load_workbook(out_buffer)
                        ws = wb.active
                        
                        green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                        orange_fill = PatternFill(start_color="FFA500", end_color="FFA500", fill_type="solid")
                        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                        
                        col_indices = {}
                        for cell in ws[1]:
                            col_indices[cell.value] = cell.column
                            
                        for row in range(2, ws.max_row + 1):
                            # Total Marks
                            tm_col = col_indices.get("Total Marks")
                            if tm_col:
                                tm_cell = ws.cell(row=row, column=tm_col)
                                try:
                                    val = float(tm_cell.value)
                                    if val > 150: tm_cell.fill = green_fill
                                    elif 135 <= val <= 150: tm_cell.fill = yellow_fill
                                    elif 120 <= val < 135: tm_cell.fill = orange_fill
                                    else: tm_cell.fill = red_fill
                                except: pass

                            # Math Marks
                            mm_col = col_indices.get("Math Marks")
                            if mm_col:
                                mm_cell = ws.cell(row=row, column=mm_col)
                                try:
                                    val = float(mm_cell.value)
                                    if val > 45: mm_cell.fill = green_fill
                                    elif 35 <= val <= 45: mm_cell.fill = yellow_fill
                                    elif 25 <= val < 35: mm_cell.fill = orange_fill
                                    else: mm_cell.fill = red_fill
                                except: pass
                                    
                            # Chemistry Marks
                            cm_col = col_indices.get("Chemistry Marks")
                            if cm_col:
                                cm_cell = ws.cell(row=row, column=cm_col)
                                try:
                                    val = float(cm_cell.value)
                                    if val > 65: cm_cell.fill = green_fill
                                    elif 50 <= val <= 65: cm_cell.fill = yellow_fill
                                    elif 40 <= val < 50: cm_cell.fill = orange_fill
                                    else: cm_cell.fill = red_fill
                                except: pass
                                    
                            # Physics Marks
                            pm_col = col_indices.get("Physics Marks")
                            if pm_col:
                                pm_cell = ws.cell(row=row, column=pm_col)
                                try:
                                    val = float(pm_cell.value)
                                    if val > 60: pm_cell.fill = green_fill
                                    elif 50 <= val <= 60: pm_cell.fill = yellow_fill
                                    elif 40 <= val < 50: pm_cell.fill = orange_fill
                                    else: pm_cell.fill = red_fill
                                except: pass

                        final_buffer = io.BytesIO()
                        wb.save(final_buffer)
                        final_buffer.seek(0)
                        
                        st.download_button(
                            label="⬇️ Download Graded Results (Excel)",
                            data=final_buffer,
                            file_name="Graded_JEE_Student_Results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    except Exception as e:
                        st.error(f"Error grading JEE DAT file: {e}")

                elif "neet" in selected_batch.lower():
                    st.info("Grading NEET student DAT file against the Answer Key...")
                    try:
                        ans_map = {'A':'1', 'B':'2', 'C':'3', 'D':'4'}
                        ak_dict = {}
                        for _, row in df_answer_key.iterrows():
                            q_num = int(row["Question Number"])
                            ans_raw = str(row["Answer"]).strip().upper()
                            ans_val = ans_map.get(ans_raw, ans_raw)
                            ak_dict[q_num] = ans_val
                            
                        results = []
                        content = dat_file.read().decode('utf-8').splitlines()
                        
                        clean_lines = [line.strip() for line in content if line.strip()]
                        
                        # 1=1, 2=2, 4=3, 8=4 mapping for NEET dat files
                        student_code_map = {'1': '1', '2': '2', '4': '3', '8': '4'}
                        
                        for i in range(0, len(clean_lines), 4):
                            if i + 3 >= len(clean_lines):
                                break
                            
                            line1 = clean_lines[i]
                            if "No.=" in line1:
                                student_id = line1.split("No.=")[-1].strip()
                            else:
                                student_id = line1.replace("No.", "").replace("=", "").strip()
                                
                            line4 = clean_lines[i+3]
                            responses_raw = line4.split()
                            
                            responses = []
                            for r in responses_raw:
                                responses.append(student_code_map.get(r, '0'))
                                
                            responses = responses[:200] + ['0'] * max(0, 200 - len(responses))
                            
                            phys_resp = responses[0:45]
                            chem_resp = responses[45:90]
                            bot_resp = responses[90:135]
                            zoo_resp = responses[135:180]
                            # 181 to 200 are ignored fillers
                            
                            def score_neet_section(resp_list, start_q):
                                marks = 0
                                wrong = 0
                                unattempted = 0
                                
                                for idx, r in enumerate(resp_list):
                                    q = start_q + idx
                                    correct_ans = ak_dict.get(q, None)
                                    
                                    if r == '0':
                                        unattempted += 1
                                    elif correct_ans and r == str(correct_ans):
                                        marks += 4
                                    else:
                                        marks -= 1
                                        wrong += 1
                                        
                                return marks, wrong, unattempted

                            p_m, p_w, p_u = score_neet_section(phys_resp, 1)
                            c_m, c_w, c_u = score_neet_section(chem_resp, 46)
                            b_m, b_w, b_u = score_neet_section(bot_resp, 91)
                            z_m, z_w, z_u = score_neet_section(zoo_resp, 136)
                            
                            total = p_m + c_m + b_m + z_m
                            
                            results.append({
                                "Student Code": student_id,
                                "Physics Marks": p_m,
                                "Physics Wrong": p_w,
                                "Physics Unattempted": p_u,
                                "Chemistry Marks": c_m,
                                "Chemistry Wrong": c_w,
                                "Chemistry Unattempted": c_u,
                                "Botany Marks": b_m,
                                "Botany Wrong": b_w,
                                "Botany Unattempted": b_u,
                                "Zoology Marks": z_m,
                                "Zoology Wrong": z_w,
                                "Zoology Unattempted": z_u,
                                "Total Marks": total
                            })
                            
                        df_results = pd.DataFrame(results)
                        if not df_results.empty:
                            df_results = df_results.sort_values(by="Total Marks", ascending=False).reset_index(drop=True)
                        
                        st.success("NEET DAT file graded successfully!")
                        st.write("### Graded Results Preview")
                        st.dataframe(df_results.head(10))
                        
                        out_buffer = io.BytesIO()
                        df_results.to_excel(out_buffer, index=False)
                        out_buffer.seek(0)
                        
                        st.download_button(
                            label="⬇️ Download Graded Results (Excel)",
                            data=out_buffer,
                            file_name="Graded_NEET_Results.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
                    except Exception as e:
                        st.error(f"Error grading NEET DAT file: {e}")
