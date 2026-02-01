import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Teacher's Marksheet Pro", layout="wide")

st.title("📊 School Exam Data Consolidator (6 Subjects)")

# Instruction section
st.markdown("""
### Steps for Teachers:
1. **Upload** the Excel file with sheets: `FIRST UNIT TEST`, `FIRST TERM`, `SECOND UNIT TEST`, `ANNUAL EXAM`.
2. **Review/Edit** the marks in the table. Enter **'AB'** for absent students.
3. **Generate Report** to calculate final results, 2-decimal percentages, and **Ranks**.
""")

def custom_round(x):
    """Standard school rounding: .5 and above goes up."""
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except:
        return 0

def clean_marks(val):
    """Treats 'AB', spaces, or empty cells as 0 for calculations."""
    if isinstance(val, str):
        v = val.strip().upper()
        if v == 'AB' or v == '':
            return 0.0
    try:
        return float(val)
    except:
        return 0.0

uploaded_file = st.file_uploader("Upload Excel Marksheet", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        
        # Define 6 Subjects
        subj_list = [f'Sub{i+1}' for i in range(6)]
        fetch_subj_map = {f'SUB{i+1}': f'Sub{i+1}' for i in range(6)}
        
        exam_configs = [
            {'label': 'FIRST UNIT TEST (25)', 'sheets': ['FIRST UNIT TEST']},
            {'label': 'FIRST TERM EXAM (50)', 'sheets': ['FIRST TERM']},
            {'label': 'SECOND UNIT TEST (25)', 'sheets': ['SECOND UNIT TEST']},
            {'label': 'ANNUAL EXAM (70/80)', 'sheets': ['ANNUAL EXAM']}
        ]

        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']
        all_students = {}

        # 1. Fetching Logic from Sheets
        for config in exam_configs:
            sheet_name = next((s for s in xl.sheet_names if s.strip().upper() in config['sheets']), None)
            
            if sheet_name:
                df = xl.parse(sheet_name)
                df.columns = df.columns.astype(str).str.strip().str.upper()
                
                # Robust column detection for Total, %, Result
                total_col = next((c for c in df.columns if 'TOTAL' in c), None)
                perc_col = next((c for c in df.columns if '%' in c or 'PERCENT' in c), None)
                res_col = next((c for c in df.columns if 'RESULT' in c), None)

                for _, row in df.iterrows():
                    roll = str(row.get('ROLL NO.', '')).strip()
                    if not roll or roll == 'nan': continue
                    
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    marks = {}
                    for k, v in fetch_subj_map.items():
                        val = row.get(k, 0)
                        # Keep 'AB' visible as text, otherwise numeric
                        marks[v] = str(val).strip() if str(val).strip().upper() == 'AB' else val

                    # Fetch existing stats from the sheet
                    marks['Grand Total'] = str(row.get(total_col, '')) if total_col else ''
                    
                    try:
                        raw_p = row.get(perc_col, '')
                        # Ensure 2-decimal formatting (Decrease Indent)
                        marks['%'] = str(round(float(raw_p), 2)) if raw_p != '' and str(raw_p).strip() != '' else str(raw_p)
                    except:
                        marks['%'] = str(row.get(perc_col, ''))

                    marks['Result'] = str(row.get(res_col, ''))
                    all_students[roll]['Exams'][config['label']] = marks

        # 2. Build Block Template (7 rows per student)
        categories = [
            'FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 
            'SECOND UNIT TEST (25)', 'ANNUAL EXAM (70/80)', 
            'INT/PRACTICAL (20/30)', 'Total Marks Out of 200',
            'Average Marks 200/2=100'
        ]

        rows = []
        # Sort students by roll number
        sorted_rolls = sorted(all_students.keys(), key=lambda x: float(x) if x.replace('.','',1).isdigit() else 0)
        
        for roll in sorted_rolls:
            s = all_students[roll]
            for cat in categories:
                row_data = {
                    'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column1': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column2': cat
                }
                for col in subj_list + result_cols: row_data[col] = ''
                
                if cat in s['Exams']:
                    row_data.update(s['Exams'][cat])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list: row_data[sub] = "0"
                
                rows.append(row_data)

        base_df = pd.DataFrame(rows)

        # Pre-format columns as strings to prevent streamlit editing errors
        for col in base_df.columns:
            base_df[col] = base_df[col].astype(str).replace('nan', '')

        # 3. Data Editor
        st.subheader("Consolidated Marksheet Editor")
        st.info("The 'Average Marks' row will be calculated and ranked after clicking the button below.")
        
        edited_df = st.data_editor(
            base_df,
            hide_index=True,
            use_container_width=True
        )

        # 4. Final Calculation and Rank Logic
        if st.button("Generate Final Report & Ranks"):
            processed_blocks = []
            pass_students_data = [] # To store (row_index_in_final_df, grand_total) for ranking

            current_row_idx = 0
            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                
                # Math Pass: Convert 'AB' to 0 for sum/avg
                numeric_marks = block.iloc[0:5][subj_list].applymap(clean_marks)
                
                # Row 5: Total 200
                total_200 = numeric_marks.sum()
                for sub in subj_list:
                    block.iloc[5, block.columns.get_loc(sub)] = str(int(total_200[sub]))
                
                # Row 6: Average 100
                avg_100 = total_200.apply(lambda x: custom_round(x/2))
                for sub in subj_list:
                    block.iloc[6, block.columns.get_loc(sub)] = str(avg_100[sub])

                # Summary Statistics for Average Row
                final_grand_total = avg_100.sum()
                final_percentage = round((final_grand_total / 600) * 100, 2)
                
                # Check Pass/Fail (35 is passing mark)
                is_pass = all(m >= 35 for m in avg_100)
                result_text = "PASS" if is_pass else "FAIL"
                
                block.iloc[6, block.columns.get_loc('Grand Total')] = str(final_grand_total)
                block.iloc[6, block.columns.get_loc('%')] = str(final_percentage)
                block.iloc[6, block.columns.get_loc('Result')] = result_text

                processed_blocks.append(block)
                
                # Rank logic: only pass students qualify
                if is_pass:
                    # Index in the final concatenated dataframe will be i+6
                    pass_students_data.append({'index': i + 6, 'total': final_grand_total})
                
            # Combine all blocks
            final_df = pd.concat(processed_blocks).reset_index(drop=True)
            
            # --- Applying Rank Logic ---
            # Sort pass students by total marks descending
            pass_students_data.sort(key=lambda x: x['total'], reverse=True)
            
            # Assign ranks (handling same marks as same rank)
            current_rank = 0
            last_total = -1
            for i, entry in enumerate(pass_students_data):
                if entry['total'] != last_total:
                    current_rank = i + 1
                last_total = entry['total']
                
                # Update the Rank column in the main dataframe
                final_df.at[entry['index'], 'Rank'] = str(current_rank)

            # 5. Excel Export
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            st.success("✅ Success! Ranks applied and percentages formatted to 56.57.")
            st.download_button("📥 Download Final Consolidated Sheet", output.getvalue(), "Final_Report.xlsx")
            
            # Show a preview of the finished data
            st.dataframe(final_df, hide_index=True)

    except Exception as e:
        st.error(f"Error encountered: {e}")
