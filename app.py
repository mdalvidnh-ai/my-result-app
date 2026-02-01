import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Result Processor Pro", layout="wide")

st.title("📊 School Exam Consolidator")
st.write("Enter Practical marks directly in the table below. The app will calculate Totals, %, Result, and Rank.")

def custom_round(x):
    # Rounds 0.5 up (e.g., 35.5 -> 36)
    return int(np.floor(x + 0.5))

uploaded_file = st.file_uploader("Upload App.xlsx", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        
        # Exact sheet names to look for
        exam_mapping = {
            'FIRST UNIT TEST': 'FIRST UNIT TEST (25)',
            'FIRST TERM': 'FIRST TERM EXAM (50)',
            'SECOND UNIT TEST': 'SECOND UNIT TEST (25)',
            'ANNUAL EXAM': 'ANNUAL EXAM (70/80)'
        }
        
        # Subject mapping (Sheet -> Template)
        subj_map = {'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'}
        subj_list = list(subj_map.values())
        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']
        all_cols = ['Roll No.', 'Column1', 'Column2'] + subj_list + result_cols

        all_students = {}

        # 1. Collect data from sheets
        for sheet_name, template_label in exam_mapping.items():
            if sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name)
                df.columns = df.columns.str.strip().str.upper() # Clean column headers
                
                for _, row in df.iterrows():
                    roll = row.get('ROLL NO.')
                    if roll is None: continue
                    
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    # Extract marks
                    marks_data = {subj_map[k]: row.get(k, 0) for k in subj_map.keys()}
                    all_students[roll]['Exams'][template_label] = marks_data

        # 2. Build the Fixed Template (7 Rows per student)
        final_rows = []
        categories = [
            'FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 
            'SECOND UNIT TEST (25)', 'ANNUAL EXAM (70/80)', 
            'INT/PRACTICAL (20/30)', 'Total Marks Out of 200',
            'Average Marks 200/2=100'
        ]

        for roll in sorted(all_students.keys()):
            s = all_students[roll]
            for cat in categories:
                row_data = {
                    'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column1': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column2': cat
                }
                # Initialize columns
                for col in subj_list + result_cols: row_data[col] = ''
                
                # Pre-fill exam marks
                if cat in s['Exams']:
                    row_data.update(s['Exams'][cat])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list: row_data[sub] = 0 # Teacher will edit this
                
                final_rows.append(row_data)

        base_df = pd.DataFrame(final_rows)

        # 3. Interactive Editor
        st.subheader("Edit Practical Marks Below")
        edited_df = st.data_editor(base_df, hide_index=True, use_container_width=True)

        if st.button("Calculate Totals & Download"):
            processed_data = []
            pass_ranking = []

            # Step through 7 rows at a time
            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                
                # Rows index 0 to 4 are the inputs (4 Exams + 1 Practical)
                # Ensure they are numeric
                for sub in subj_list:
                    block[sub] = pd.to_numeric(block[sub], errors='coerce').fillna(0)

                # Row 5: Total Out of 200
                total_200_vals = block.iloc[0:5][subj_list].sum()
                for sub in subj_list:
                    block.at[block.index[5], sub] = total_200_vals[sub]
                
                # Row 6: Average Marks 100
                avg_100_vals = total_200_vals.apply(lambda x: custom_round(x/2))
                for sub in subj_list:
                    block.at[block.index[6], sub] = avg_100_vals[sub]

                # Result & Stats for Average Row
                grand_total = avg_100_vals.sum()
                percentage = round((grand_total / 600) * 100, 2)
                is_pass = all(m >= 35 for m in avg_100_vals)
                
                block.at[block.index[6], 'Grand Total'] = grand_total
                block.at[block.index[6], '%'] = percentage
                block.at[block.index[6], 'Result'] = "PASS" if is_pass else "FAIL"

                processed_data.append(block)
                
                if is_pass:
                    pass_ranking.append({'row_idx': (i + 6), 'total': grand_total})

            # Re-assemble and apply Rank
            final_df = pd.concat(processed_data)
            pass_ranking.sort(key=lambda x: x['total'], reverse=True)
            
            # Create a dictionary for quick rank lookup
            ranks = {item['row_idx']: r+1 for r, item in enumerate(pass_ranking)}
            
            # Apply rank to the Average Marks rows
            final_rows_list = final_df.to_dict('records')
            for idx, r in enumerate(final_rows_list):
                if idx in ranks:
                    r['Rank'] = ranks[idx]

            # Export
            output_df = pd.DataFrame(final_rows_list)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            st.success("Calculations Finished!")
            st.download_button("📥 Download Excel", output.getvalue(), "Final_Consolidated.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
