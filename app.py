import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Result Processor Pro", layout="wide")

st.title("🏫 Student Exam Data Consolidator")
st.markdown("""
1. **Upload** your master Excel file.
2. **Enter** Internal/Practical marks in the table below (Row 5 of each student).
3. **Click** the button to calculate Totals, %, Result, and Rank.
""")

def custom_round(x):
    """Rounds x.5 and above up (e.g., 34.5 -> 35)."""
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except:
        return 0

uploaded_file = st.file_uploader("Upload your Excel file (App.xlsx)", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        
        # We look for these sheets. We use lowercase and strip for safety.
        available_sheets = {s.strip().upper(): s for s in xl.sheet_names}
        
        # Map Internal Labels to Sheet Names
        # If sheet 'FIRST TERM' is not found, it tries 'FIRST TERM EXAM' etc.
        def find_sheet(names):
            for n in names:
                if n.upper() in available_sheets:
                    return available_sheets[n.upper()]
            return None

        exam_configs = [
            {'label': 'FIRST UNIT TEST (25)', 'sheets': ['FIRST UNIT TEST']},
            {'label': 'FIRST TERM EXAM (50)', 'sheets': ['FIRST TERM', 'FIRST TERM EXAM']},
            {'label': 'SECOND UNIT TEST (25)', 'sheets': ['SECOND UNIT TEST']},
            {'label': 'ANNUAL EXAM (70/80)', 'sheets': ['ANNUAL EXAM']}
        ]

        subj_map = {'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'}
        subj_list = list(subj_map.values())
        all_cols = ['Roll No.', 'Column1', 'Column2'] + subj_list + ['Grand Total', '%', 'Result', 'Remark', 'Rank']

        all_students = {}

        # 1. Fetch data from all sheets
        for config in exam_configs:
            sheet_name = find_sheet(config['sheets'])
            if sheet_name:
                df = xl.parse(sheet_name)
                df.columns = df.columns.str.strip().str.upper()
                for _, row in df.iterrows():
                    roll = str(row.get('ROLL NO.', '')).strip()
                    if not roll or roll == 'nan': continue
                    
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    marks = {subj_map[k]: row.get(k, 0) for k in subj_map.keys() if k in df.columns}
                    all_students[roll]['Exams'][config['label']] = marks

        # 2. Build the Fixed Template (7 Rows per student)
        categories = [
            'FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 
            'SECOND UNIT TEST (25)', 'ANNUAL EXAM (70/80)', 
            'INT/PRACTICAL (20/30)', 'Total Marks Out of 200',
            'Average Marks 200/2=100'
        ]

        rows = []
        for roll in sorted(all_students.keys(), key=lambda x: float(x) if x.replace('.','',1).isdigit() else 0):
            s = all_students[roll]
            for cat in categories:
                row_data = {
                    'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column1': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                    'Column2': cat
                }
                # Init other columns
                for col in subj_list + ['Grand Total', '%', 'Result', 'Remark', 'Rank']:
                    row_data[col] = ''
                
                # Fill Exam data
                if cat in s['Exams']:
                    row_data.update(s['Exams'][cat])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list: row_data[sub] = 0 # Teacher entry
                
                rows.append(row_data)

        base_df = pd.DataFrame(rows)

        # 3. Interactive Data Editor
        st.subheader("Edit/Review Data")
        st.info("💡 Enter Internal/Practical marks in the row: 'INT/PRACTICAL (20/30)'")
        
        # Ensure subject columns are numeric for the editor
        for sub in subj_list:
            base_df[sub] = pd.to_numeric(base_df[sub], errors='coerce').fillna(0)

        edited_df = st.data_editor(
            base_df,
            column_config={
                "Roll No.": st.column_config.Column(disabled=True),
                "Column1": st.column_config.Column("Student Name", disabled=True),
                "Column2": st.column_config.Column("Exam Type", disabled=True),
            },
            hide_index=True,
            use_container_width=True
        )

        # 4. Final Calculation logic
        if st.button("Calculate Final Summary & Download"):
            processed_data = []
            pass_ranking = []

            # Step through each student block (7 rows)
            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                
                # We need to sum rows 0, 1, 2, 3 (Exams) and 4 (Practical)
                # Ensure they are floats for summing
                input_data = block.iloc[0:5][subj_list].apply(pd.to_numeric, errors='coerce').fillna(0)
                
                # Row index 5: Total Out of 200
                total_200 = input_data.sum()
                for sub in subj_list:
                    block.iloc[5, block.columns.get_loc(sub)] = total_200[sub]
                
                # Row index 6: Average Marks 100
                avg_100 = total_200.apply(lambda x: custom_round(x/2))
                for sub in subj_list:
                    block.iloc[6, block.columns.get_loc(sub)] = avg_100[sub]

                # Result & Statistics for the Average Row (Row 6)
                grand_total = avg_100.sum()
                perc = round((grand_total / 600) * 100, 2)
                # Pass logic: Every subject must be >= 35
                is_pass = all(m >= 35 for m in avg_100)
                
                block.iloc[6, block.columns.get_loc('Grand Total')] = grand_total
                block.iloc[6, block.columns.get_loc('%')] = perc
                block.iloc[6, block.columns.get_loc('Result')] = "PASS" if is_pass else "FAIL"

                processed_data.append(block)
                
                if is_pass:
                    pass_ranking.append({'row_idx': (i + 6), 'total': grand_total})

            # Final Rank Logic
            final_df = pd.concat(processed_data)
            pass_ranking.sort(key=lambda x: x['total'], reverse=True)
            
            # Create a dictionary for mapping row index to rank
            ranks_map = {item['row_idx']: r+1 for r, item in enumerate(pass_ranking)}
            
            # Apply Ranks to the combined dataframe
            final_records = final_df.to_dict('records')
            for idx, r in enumerate(final_records):
                if idx in ranks_map:
                    r['Rank'] = ranks_map[idx]

            # 5. Export to Excel
            output_df = pd.DataFrame(final_records)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            st.success("✅ Calculations Complete!")
            st.download_button("📥 Download Consolidated Results", output.getvalue(), "Consolidated_Results.xlsx")

    except Exception as e:
        st.error(f"An error occurred: {e}")
