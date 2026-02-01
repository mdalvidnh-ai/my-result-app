import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Result Processor Pro", layout="wide")

st.title("🏫 Student Exam Data Consolidator")
st.markdown("### Instructions:")
st.markdown("""
- **Upload** the Excel file with sheets: 'FIRST UNIT TEST', 'FIRST TERM', 'SECOND UNIT TEST', 'ANNUAL EXAM'.
- **Column Headers** in sheets must include: `ROLL NO.`, `STUDENT NAME`, `SUB1`, `SUB2`, `SUB3`, `SUB4`, `SUB5`, `SUB6`, `SUB7`, `TOTAL`, `%`, `RESULT`.
- **'AB'** is treated as 0 in totals but stays visible as 'AB'.
""")

def custom_round(x):
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except:
        return 0

def clean_marks(val):
    if isinstance(val, str):
        v = val.strip().upper()
        if v == 'AB' or v == '':
            return 0.0
    try:
        return float(val)
    except:
        return 0.0

def highlight_avg_row(row):
    if row['Column2'] == 'Average Marks 200/2=100':
        return ['background-color: #ffffcc'] * len(row)
    return [''] * len(row)

uploaded_file = st.file_uploader("Upload Master Excel File", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        
        # Mapping generic subject names
        subj_list = ['Sub1', 'Sub2', 'Sub3', 'Sub4', 'Sub5', 'Sub6', 'Sub7']
        # Internal mapping for fetching (looks for SUB1, SUB2... in excel)
        fetch_subj_map = {f'SUB{i+1}': f'Sub{i+1}' for i in range(7)}
        
        exam_configs = [
            {'label': 'FIRST UNIT TEST (25)', 'sheets': ['FIRST UNIT TEST']},
            {'label': 'FIRST TERM EXAM (50)', 'sheets': ['FIRST TERM']},
            {'label': 'SECOND UNIT TEST (25)', 'sheets': ['SECOND UNIT TEST']},
            {'label': 'ANNUAL EXAM (70/80)', 'sheets': ['ANNUAL EXAM']}
        ]

        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']
        all_students = {}

        # 1. Fetching Logic
        for config in exam_configs:
            sheet_name = next((s for s in xl.sheet_names if s.strip().upper() in config['sheets']), None)
            
            if sheet_name:
                df = xl.parse(sheet_name)
                df.columns = df.columns.astype(str).str.strip().str.upper()
                
                # Dynamic column identification
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
                        # Keep AB as string, otherwise keep numeric
                        marks[v] = str(val).strip() if str(val).strip().upper() == 'AB' else val

                    # Fetch Grand Total, % and Result
                    marks['Grand Total'] = str(row.get(total_col, '')) if total_col else ''
                    
                    try:
                        raw_p = row.get(perc_col, '')
                        # Rounding % to 2 decimals for "Decrease Indent" requirement
                        marks['%'] = str(round(float(raw_p), 2)) if raw_p != '' and str(raw_p).strip() != '' else str(raw_p)
                    except:
                        marks['%'] = str(row.get(perc_col, ''))

                    marks['Result'] = str(row.get(res_col, ''))
                    all_students[roll]['Exams'][config['label']] = marks

        # 2. 7-Row Template Construction
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
                for col in subj_list + result_cols: row_data[col] = ''
                
                if cat in s['Exams']:
                    row_data.update(s['Exams'][cat])
                elif cat == 'INT/PRACTICAL (20/30)':
                    for sub in subj_list: row_data[sub] = "0"
                
                rows.append(row_data)

        base_df = pd.DataFrame(rows)

        # Force all columns to string to avoid "Invalid value for dtype string" crashes
        for col in base_df.columns:
            base_df[col] = base_df[col].astype(str).replace('nan', '')

        # 3. Data Editor with Row Styling
        st.subheader("Edit/Review Consolidated Marks")
        
        # Displaying the editor
        edited_df = st.data_editor(
            base_df,
            hide_index=True,
            use_container_width=True
        )

        # 4. Calculation Block
        if st.button("Generate Final Report"):
            processed_data = []
            pass_ranking = []

            for i in range(0, len(edited_df), 7):
                block = edited_df.iloc[i:i+7].copy()
                # Use clean_marks to treat 'AB' as 0
                numeric_marks = block.iloc[0:5][subj_list].applymap(clean_marks)
                
                # Calculate 'Total Marks Out of 200'
                total_200 = numeric_marks.sum()
                for sub in subj_list:
                    block.iloc[5, block.columns.get_loc(sub)] = str(int(total_200[sub]))
                
                # Calculate 'Average Marks 100'
                avg_100 = total_200.apply(lambda x: custom_round(x/2))
                for sub in subj_list:
                    block.iloc[6, block.columns.get_loc(sub)] = str(avg_100[sub])

                # Final Row (Average) Totals and Ranking
                g_total = avg_100.sum()
                perc = round((g_total / 700) * 100, 2) # Updated to 700 for 7 subjects
                is_pass = all(m >= 35 for m in avg_100)
                
                block.iloc[6, block.columns.get_loc('Grand Total')] = str(g_total)
                block.iloc[6, block.columns.get_loc('%')] = str(perc)
                block.iloc[6, block.columns.get_loc('Result')] = "PASS" if is_pass else "FAIL"

                processed_data.append(block)
                if is_pass:
                    pass_ranking.append({'row_idx': (i + 6), 'total': g_total})

            # Calculate Global Rank for PASS students
            final_df = pd.concat(processed_data)
            pass_ranking.sort(key=lambda x: x['total'], reverse=True)
            ranks_map = {item['row_idx']: r+1 for r, item in enumerate(pass_ranking)}
            
            final_records = final_df.to_dict('records')
            for idx, r in enumerate(final_records):
                if idx in ranks_map:
                    r['Rank'] = str(ranks_map[idx])

            # 5. Export to Excel
            output_df = pd.DataFrame(final_records)
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                output_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            st.success("✅ Success! Total fetched and Percentage rounded (e.g., 56.57).")
            st.download_button("📥 Download Final Report", output.getvalue(), "Final_Consolidated.xlsx")

    except Exception as e:
        st.error(f"Something went wrong: {e}")
