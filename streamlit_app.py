import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
import io
import os

st.set_page_config(page_title="Student Excel Analyzer", page_icon="📊", layout="centered")

st.title("Student Results Analyzer 📊")
st.write("Upload your weekly Aakash Excel file. This app will instantly filter the NSPIRA-CC branch, format the ranks (Top 33% = Green, Middle = Yellow, Bottom = Red), and give you the clean file!")

uploaded_file = st.file_uploader("Upload Student Results Excel", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    st.info("Reading file...")
    
    try:
        # Read without headers to pick columns by dynamic search
        if uploaded_file.name.lower().endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file, header=None)
        else:
            df_raw = pd.read_excel(uploaded_file, header=None)
        
        # Check if file has enough rows
        if len(df_raw) < 3:
            st.error("Error: Excel file does not contain enough rows to analyze.")
            st.stop()
            
        data_rows = df_raw.iloc[2:]
        
        # Dynamic Column Discovery
        header_row = [str(x).strip() for x in df_raw.iloc[1].tolist()]

        def get_first_idx(col_name_or_list, start_idx=0):
            col_names = col_name_or_list if isinstance(col_name_or_list, list) else [col_name_or_list]
            for i in range(start_idx, len(header_row)):
                if header_row[i] in col_names:
                    return i
            st.error(f"Error: Could not find column '{col_names}' in the Excel header.")
            st.stop()
            
        branch_idx = get_first_idx('BRANCH')
        name_idx = get_first_idx('NAME')
        
        mm_idx = get_first_idx('MM')
        mr_idx = get_first_idx('MR', mm_idx)
        mw_idx = get_first_idx('W', mm_idx)
        
        pm_idx = get_first_idx('PM', mw_idx)
        pr_idx = get_first_idx('PR', pm_idx)
        pw_idx = get_first_idx('W', pm_idx)
        
        cm_idx = get_first_idx('CM', pw_idx)
        cr_idx = get_first_idx('CR', cm_idx)
        cw_idx = get_first_idx('W', cm_idx)
        
        tm_idx = get_first_idx(['TM', 'Total'], cw_idx)
        tr_idx = get_first_idx('TR', tm_idx)

        # Allow user to specify branch
        target_branch = st.text_input("Enter branch to filter by:", value="MH-MUM-NSP-NSPIRA-CC")

        st.info(f"Filtering for branch {target_branch}...")
        df_filtered = data_rows[data_rows[branch_idx] == target_branch].copy()

        if df_filtered.empty:
            st.warning(f"No students found for branch '{target_branch}'.")
            st.stop()

        req_indices = [name_idx, mm_idx, mr_idx, mw_idx, pm_idx, pr_idx, pw_idx, cm_idx, cr_idx, cw_idx, tm_idx, tr_idx]
        df_final = df_filtered.iloc[:, req_indices].copy()

        # Set headers
        df_final.columns = ['NAME', 'MM', 'MR', 'Math W', 'PM', 'PR', 'Physics W', 'CM', 'CR', 'Chemistry W', 'TM', 'TR']
        df_final = df_final.reset_index(drop=True)

        cols_to_convert = ['MM', 'MR', 'Math W', 'PM', 'PR', 'Physics W', 'CM', 'CR', 'Chemistry W', 'TM', 'TR']
        for col in cols_to_convert:
            df_final[col] = pd.to_numeric(df_final[col], errors='coerce')

        def get_tertiles(series):
            s = series.dropna()
            if len(s) == 0: return [], [], []
            try:
                bins = pd.qcut(s, q=3, labels=False, duplicates='drop')
                best_indices = bins[bins == 0].index.tolist()
                avg_indices = bins[bins == 1].index.tolist()
                worst_indices = bins[bins == 2].index.tolist()
                
                if len(set(bins)) < 3:
                     p33 = s.quantile(0.33)
                     p66 = s.quantile(0.66)
                     best_indices = s[s <= p33].index.tolist()
                     avg_indices = s[(s > p33) & (s <= p66)].index.tolist()
                     worst_indices = s[s > p66].index.tolist()
            except:
                return [], [], []
            return best_indices, worst_indices, avg_indices

        st.info("Calculating Rank percentiles and coloring...")
        best_m, worst_m, avg_m = get_tertiles(df_final['MR'])
        best_p, worst_p, avg_p = get_tertiles(df_final['PR'])
        best_c, worst_c, avg_c = get_tertiles(df_final['CR'])
        best_t, worst_t, avg_t = get_tertiles(df_final['TR'])

        # Save to memory buffer instead of disk for web download 
        output_buffer = io.BytesIO()
        df_final.to_excel(output_buffer, index=False)
        output_buffer.seek(0)
        
        # Apply colors using openpyxl on the buffer
        wb = openpyxl.load_workbook(output_buffer)
        ws = wb.active

        green_fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        def apply_color(ws, col_name, best_indices, worst_indices, avg_indices):
            col_idx = None
            for cell in ws[1]:
                if cell.value == col_name:
                    col_idx = cell.column
                    break
            if not col_idx: return
            
            for idx in best_indices: ws.cell(row=idx+2, column=col_idx).fill = green_fill
            for idx in avg_indices: ws.cell(row=idx+2, column=col_idx).fill = yellow_fill
            for idx in worst_indices: ws.cell(row=idx+2, column=col_idx).fill = red_fill

        apply_color(ws, 'MR', best_m, worst_m, avg_m)
        apply_color(ws, 'PR', best_p, worst_p, avg_p)
        apply_color(ws, 'CR', best_c, worst_c, avg_c)
        apply_color(ws, 'TR', best_t, worst_t, avg_t)

        final_buffer = io.BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)
        
        st.success("Analysis Complete! Your file is ready to download.")
        
        # Display a quick preview table
        st.write("### Data Preview (First 5 Rows)")
        st.dataframe(df_final.head())
        
        # Provide the download button
        base_name, _ = os.path.splitext(uploaded_file.name)
        new_filename = f"{base_name}_Analyzed.xlsx"
        st.download_button(
            label="⬇️ Download Formatted Excel File",
            data=final_buffer,
            file_name=new_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"An error occurred while processing the file: {e}")
