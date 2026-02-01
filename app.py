import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Consolidated Marksheet Pro", layout="wide")

st.title("📊 School Exam Consolidator")
st.info("Structure: 7 Rows per Student | 13 Columns Total")

def custom_round(x):
    # Rule: 0.5 and above rounds up (e.g., 34.5 -> 35)
    return int(np.floor(x + 0.5))

uploaded_file = st.file_uploader("Upload App.xlsx", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        exam_names = ['FIRST UNIT TEST', 'FIRST TERM', 'SECOND UNIT TEST', 'ANNUAL EXAM']
        
        # Exact column names from your template
        subj_list = ['Eng', 'Mar', 'Geo', 'Soc', 'Psy', 'Eco']
        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']
        all_cols = ['Roll No.', 'Column1', 'Column2'] + subj_list + result_cols

        # Mapping for data extraction
        subj_map = {'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'}

        all_students = {}

        # 1. Collect data
        for sheet in exam_names:
            if sheet in xl.sheet_names:
                df = xl.parse(sheet)
                df.columns = df.columns.str.strip().str.upper()
                for _, row in df.iterrows():
                    roll = row['ROLL NO.']
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    # Pull marks and sheet-specific results
                    data = {subj_map[k]: row.get(k, 0) for k in subj_map.keys()}
                    data['Grand Total'] = row.get('TOTAL', '')
                    data['%'] = row.get('%', '')
                    data['Result'] = row.get('RESULT', '')
                    all_students[roll]['Exams'][sheet] = data

        # 2. Build the Fixed Template (7 Rows per student)
        final_rows = []
        for roll in sorted(all_students.keys()):
            s = all_students[roll]
            categories = [
                'FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 
                'SECOND UNIT TEST (25)', 'ANNUAL EXAM (70/80)', 
                'INT/PRACTICAL (20/30)', 'Total Marks Out of 200',
                'Average Marks 200/2=100'
            ]
            
            for cat in categories:
                row_data = {
                    'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column1': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column2': cat
                }
                # Initialize all columns as empty
                for col in subj_list + result_cols: row_data[col] = ''
                
                # Fill data for the 4 exam rows
                key = cat.split(' (')[0]
                if key in s['Exams']:
                    row_data.update(s['Exams'][key])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list: row_data[sub] = 0 # Default for teacher entry
                
                final_rows.append(row_data)

        base_df = pd.DataFrame(final_rows)

        # 3. Editable Table
        st.subheader("Interactive Marksheet")
        st.write("Enter Practical marks in the 5th row of each student block.")
        edited_df = st.data_editor(base_df, hide_index=True, use_container_width=True)

        if st.button("Finalize All Calculations & Ranks"):
            processed_data = []
            pass_list = [] # For ranking

            # Process in blocks of 7
            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                
                # Rows 0-4 are inputs (Exams + Practical)
                input_rows = block.iloc[0:5]
                
                # Row 5: Total Out of 200
                total_row = block.iloc[5].to_dict()
                # Row 6: Average Marks
                avg_row = block.iloc[6].to_dict()
                
                subjects_for_pass_check = []
                grand_total_100 = 0

                for sub in subj_list:
                    # Sum the first 5 rows for this subject
                    s_sum = pd.to_numeric(input_rows[sub], errors='coerce').sum()
                    total_row[sub] = s_sum
                    
                    # Round average
                    rounded_avg = custom_round(s_sum / 2)
                    avg_row[sub] = rounded_avg
                    
                    subjects_for_pass_check.append(rounded_avg)
                    grand_total_100 += rounded_avg

                # Average Row Calculations
                avg_row['Grand Total'] = grand_total_100
                avg_row['%'] = round((grand_total_100 / 600) * 100, 2)
                
                is_pass = all(m >= 35 for m in subjects_for_pass_check)
                avg_row['Result'] = "PASS" if is_pass else "FAIL"

                # Re-add all 7 rows to the final list
                for j in range(5):
                    processed_data.append(block.iloc[j].to_dict())
                processed_data.append(total_row)
                processed_data.append(avg_row)

                # Store index for ranking
                if is_pass:
                    pass_list.append({'index': len(processed_data)-1, 'total': grand_total_100})

            # 4. Global Ranking for Pass students
            pass_list.sort(key=lambda x: x['total'], reverse=True)
            for rank, item in enumerate(pass_list, 1):
                processed_data[item['index']]['Rank'] = rank

            # 5. Export
            output_df = pd.DataFrame(processed_data)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            st.success("Calculated! Only PASS students are ranked based on Average Marks row.")
            st.download_button("📥 Download Excel", output.getvalue(), "Final_Consolidated.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
