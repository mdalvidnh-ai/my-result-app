import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Result Processor Pro", layout="wide")

st.title("🏫 Student Exam Data Consolidator")
st.markdown("### Instructions:")
st.markdown("""
- **Upload** the provided 'Teacher_Template.xlsx'.
- **Enter** Internal/Practical marks in the table.
- **'AB'** is treated as 0 in totals but stays as 'AB' in the table.
""")

def custom_round(x):
    """Rounds x.5 and above up (e.g., 34.5 -> 35)."""
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except:
        return 0

def clean_marks(val):
    """Converts AB or empty to 0 for math."""
    if isinstance(val, str):
        v = val.strip().upper()
        if v == 'AB' or v == '':
            return 0.0
    try:
        return float(val)
    except:
        return 0.0

def highlight_avg_row(row):
    """Styles the Average row with a light yellow background."""
    if row['Column2'] == 'Average Marks 200/2=100':
        return ['background-color: #ffffcc'] * len(row)
    return [''] * len(row)

uploaded_file = st.file_uploader("Upload Excel File", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        
        # Exact sheet search
        exam_configs = [
            {'label': 'FIRST UNIT TEST (25)', 'sheets': ['FIRST UNIT TEST']},
            {'label': 'FIRST TERM EXAM (50)', 'sheets': ['FIRST TERM', 'FIRST TERM EXAM']},
            {'label': 'SECOND UNIT TEST (25)', 'sheets': ['SECOND UNIT TEST']},
            {'label': 'ANNUAL EXAM (70/80)', 'sheets': ['ANNUAL EXAM']}
        ]

        subj_map = {'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'}
        subj_list = list(subj_map.values())
        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']

        all_students = {}

        # 1. Robust Data Fetching
        for config in exam_configs:
            # Find the sheet regardless of extra spaces
            sheet_name = next((s for s in xl.sheet_names if s.strip().upper() in config['sheets']), None)
            
            if sheet_name:
                df = xl.parse(sheet_name)
                df.columns = df.columns.astype(str).str.strip().str.upper()
                
                # Identify Total and Percent columns dynamically
                total_col = next((c for c in df.columns if 'TOTAL' in c), None)
                perc_col = next((c for c in df.columns if '%' in c or 'PERCENT' in c), None)
                res_col = next((c for c in df.columns if 'RESULT' in c), None)

                for _, row in df.iterrows():
                    roll = str(row.get('ROLL NO.', '')).strip()
                    if not roll or roll == 'nan': continue
                    
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    marks = {}
                    for k, v in subj_map.items():
                        val = row.get(k, 0)
                        marks[v] = str(val).strip() if str(val).strip().upper() == 'AB' else val

                    # Fetch Total, %, and Result safely
                    marks['Grand Total'] = str(row.get(total_col, '')) if total_col else ''
                    
                    try:
                        raw_p = row.get(perc_col, '')
                        marks['%'] = str(round(float(raw_p), 2)) if raw_p != '' and str(raw_p).strip() != '' else str(raw_p)
                    except:
                        marks['%'] = str(row.get(perc_col, ''))

                    marks['Result'] = str(row.get(res_col, ''))
                    all_students[roll]['Exams'][config['label']] = marks

        # 2. Build 7-Row Student Blocks
        categories = ['FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 'SECOND UNIT TEST (25)', 
                      'ANNUAL EXAM (70/80)', 'INT/PRACTICAL (20/30)', 'Total Marks Out of 200', 
                      'Average Marks 200/2=100']

        rows = []
        for roll in sorted(all_students.keys(), key=lambda x: float(x) if x.replace('.','',1).isdigit() else 0):
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

        # Force String type for EVERYTHING to prevent streamlit crashes
        for col in base_df.columns:
            base_df[col] = base_df[col].astype(str).replace('nan', '')

        # 3. Highlighted Data Editor
        st.subheader("Final Consolidation Table")
        
        # Display the editor
        # Note: Background colors only apply to static tables. 
        # For the editor, we use the raw dataframe.
        edited_df = st.data_editor(
            base_df,
            hide_index=True,
            use_container_width=True
        )

        # 4. Calculation Logic
        if st.button("Generate Final Report"):
            processed_data = []
            pass_ranking = []

            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                # Use clean_marks to treat AB as 0
                numeric_marks = block.iloc[0:5][subj_list].applymap(clean_marks)
                
                # Total 200 row
                total_200 = numeric_marks.sum()
                for sub in subj_list:
                    block.iloc[5, block.columns.get_loc(sub)] = str(int(total_200[sub]))
                
                # Average 100 row
                avg_100 = total_200.apply(lambda x: custom_round(x/2))
                for sub in subj_list:
                    block.iloc[6, block.columns.get_loc(sub)] = str(avg_100[sub])

                # Grand Stats for Average Row
                g_total = avg_100.sum()
                perc = round((g_total / 600) * 100, 2)
                is_pass = all(m >= 35 for m in avg_100)
                
                block.iloc[6, block.columns.get_loc('Grand Total')] = str(g_total)
                block.iloc[6, block.columns.get_loc('%')] = str(perc)
                block.iloc[6, block.columns.get_loc('Result')] = "PASS" if is_pass else "FAIL"

                processed_data.append(block)
                if is_pass:
                    pass_ranking.append({'row_idx': (i + 6), 'total': g_total})

            # Apply Rank to PASS students
            final_df = pd.concat(processed_data)
            pass_ranking.sort(key=lambda x: x['total'], reverse=True)
            ranks_map = {item['row_idx']: r+1 for r, item in enumerate(pass_ranking)}
            
            final_records = final_df.to_dict('records')
            for idx, r in enumerate(final_records):
                if idx in ranks_map:
                    r['Rank'] = str(ranks_map[idx])

            # 5. Export
            output_df = pd.DataFrame(final_records)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            st.success("✅ Calculations finished. Totals fetched and Percentage rounded.")
            st.download_button("📥 Download Final Report", output.getvalue(), "Final_Consolidated_Results.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
