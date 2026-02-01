import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Result Processor Pro", layout="wide")

st.title("🏫 Student Exam Data Consolidator")

def custom_round(x):
    """Rounds x.5 and above up (e.g., 34.5 -> 35)."""
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except:
        return 0

def clean_marks(val):
    """Converts AB/ab or strings to 0 for math calculations."""
    if isinstance(val, str):
        if val.strip().upper() == 'AB':
            return 0.0
    try:
        return float(val)
    except:
        return 0.0

uploaded_file = st.file_uploader("Upload App.xlsx", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        available_sheets = {s.strip().upper(): s for s in xl.sheet_names}
        
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
        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']

        all_students = {}

        # 1. Fetch data from sheets
        for config in exam_configs:
            sheet_name = find_sheet(config['sheets'])
            if sheet_name:
                df = xl.parse(sheet_name)
                df.columns = df.columns.astype(str).str.strip().str.upper()
                
                for _, row in df.iterrows():
                    roll = str(row.get('ROLL NO.', '')).strip()
                    if not roll or roll == 'nan': continue
                    
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    marks = {}
                    for k, v in subj_map.items():
                        val = row.get(k, 0)
                        marks[v] = str(val).strip() if str(val).strip().upper() == 'AB' else val

                    # Fetch Total, %, and Result
                    # Using str() and rounding immediately to prevent float errors
                    val_total = row.get('TOTAL', row.get('GRAND TOTAL', ''))
                    marks['Grand Total'] = str(val_total) if val_total != '' else ''
                    
                    raw_perc = row.get('%', row.get('PERCENTAGE', ''))
                    try:
                        # Decrease Indent: Rounding to 2 decimal places here
                        marks['%'] = str(round(float(raw_perc), 2)) if raw_perc != '' else ''
                    except:
                        marks['%'] = str(raw_perc)
                        
                    marks['Result'] = str(row.get('RESULT', ''))
                    all_students[roll]['Exams'][config['label']] = marks

        # 2. Build the Fixed Template (7 Rows)
        categories = ['FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 'SECOND UNIT TEST (25)', 
                      'ANNUAL EXAM (70/80)', 'INT/PRACTICAL (20/30)', 'Total Marks Out of 200', 
                      'Average Marks 200/2=100']

        rows = []
        for roll in sorted(all_students.keys(), key=lambda x: float(x) if x.replace('.','',1).isdigit() else 0):
            s = all_students[roll]
            for cat in categories:
                row_data = {'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                            'Column1': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                            'Column2': cat}
                for col in subj_list + result_cols: row_data[col] = ''
                
                if cat in s['Exams']:
                    row_data.update(s['Exams'][cat])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list: row_data[sub] = 0
                
                rows.append(row_data)

        base_df = pd.DataFrame(rows)

        # CRITICAL FIX: Force every column to be a String before the editor opens
        # This prevents the "float64 vs string" crash
        for col in base_df.columns:
            base_df[col] = base_df[col].astype(str).replace('nan', '')

        # 3. Data Editor
        st.subheader("Edit Practical Marks & Review Sheet Data")
        st.info("AB/ab is treated as 0 for calculations. Percentage is rounded to 2 decimals.")
        
        edited_df = st.data_editor(
            base_df,
            hide_index=True,
            use_container_width=True
        )

        # 4. Final Calculation
        if st.button("Generate Final Consolidated Report"):
            processed_data = []
            pass_ranking = []

            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                
                # Math conversion for first 5 rows
                numeric_marks = block.iloc[0:5][subj_list].applymap(clean_marks)
                
                # Row 5: Total 200
                total_200 = numeric_marks.sum()
                for sub in subj_list:
                    block.iloc[5, block.columns.get_loc(sub)] = str(int(total_200[sub]))
                
                # Row 6: Average 100
                avg_100 = total_200.apply(lambda x: custom_round(x/2))
                for sub in subj_list:
                    block.iloc[6, block.columns.get_loc(sub)] = str(avg_100[sub])

                # Final Row Statistics (Average Marks Row)
                grand_total_final = avg_100.sum()
                # Decrease Indent: Round to 2 decimals (e.g., 58.17)
                perc_final = round((grand_total_final / 600) * 100, 2)
                is_pass = all(m >= 35 for m in avg_100)
                
                block.iloc[6, block.columns.get_loc('Grand Total')] = str(grand_total_final)
                block.iloc[6, block.columns.get_loc('%')] = str(perc_final)
                block.iloc[6, block.columns.get_loc('Result')] = "PASS" if is_pass else "FAIL"

                processed_data.append(block)
                if is_pass:
                    pass_ranking.append({'row_idx': (i + 6), 'total': grand_total_final})

            # Apply Rank
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
            
            st.success("✅ Success! Fixed data-type error and rounded percentages.")
            st.download_button("📥 Download Final Sheet", output.getvalue(), "Final_Consolidated.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
