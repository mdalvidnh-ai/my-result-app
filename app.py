import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Teacher's Marksheet Pro", layout="wide")

st.title("📊 School Exam Data Consolidator")
st.markdown("### Instructions for Teachers:")
st.markdown("""
1. Upload the Excel file containing the 4 exam sheets.
2. Enter **Internal/Practical** marks directly in the table below.
3. Use **'AB'** for absent students (it will be calculated as 0).
4. Click **Generate Report** to calculate final totals, percentages, and ranks.
""")

def custom_round(x):
    try:
        val = float(x)
        return int(np.floor(val + 0.5))
    except:
        return 0

def clean_marks(val):
    if isinstance(val, str):
        if val.strip().upper() == 'AB':
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
        
        # Exact sheet names expected
        exam_configs = [
            {'label': 'FIRST UNIT TEST (25)', 'sheet': 'FIRST UNIT TEST'},
            {'label': 'FIRST TERM EXAM (50)', 'sheet': 'FIRST TERM'},
            {'label': 'SECOND UNIT TEST (25)', 'sheet': 'SECOND UNIT TEST'},
            {'label': 'ANNUAL EXAM (70/80)', 'sheet': 'ANNUAL EXAM'}
        ]

        # Exact column headers expected in those sheets
        subj_map = {'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'}
        subj_list = list(subj_map.values())
        result_cols = ['Grand Total', '%', 'Result', 'Remark', 'Rank']

        all_students = {}

        # 1. Fetch data from sheets
        for config in exam_configs:
            if config['sheet'] in xl.sheet_names:
                df = xl.parse(config['sheet'])
                df.columns = df.columns.astype(str).str.strip().str.upper()
                
                for _, row in df.iterrows():
                    roll = str(row.get('ROLL NO.', '')).strip()
                    if not roll or roll == 'nan': continue
                    
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    marks = {v: (str(row.get(k, 0)).strip() if str(row.get(k, 0)).strip().upper() == 'AB' else row.get(k, 0)) for k, v in subj_map.items()}
                    
                    # Fetching Sheet Totals
                    val_total = row.get('TOTAL', '')
                    marks['Grand Total'] = str(val_total)
                    
                    raw_perc = row.get('%', '')
                    try:
                        marks['%'] = str(round(float(raw_perc), 2)) if raw_perc != '' else ''
                    except:
                        marks['%'] = str(raw_perc)
                        
                    marks['Result'] = str(row.get('RESULT', ''))
                    all_students[roll]['Exams'][config['label']] = marks

        # 2. Build 7-Row Template
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
        for col in base_df.columns:
            base_df[col] = base_df[col].astype(str).replace('nan', '')

        # 3. Interactive Data Editor
        st.subheader("Student Marks Table")
        styled_view = base_df.style.apply(highlight_avg_row, axis=1)
        
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
                numeric_marks = block.iloc[0:5][subj_list].applymap(clean_marks)
                
                # Row 5: Total 200
                total_200 = numeric_marks.sum()
                for sub in subj_list:
                    block.iloc[5, block.columns.get_loc(sub)] = str(int(total_200[sub]))
                
                # Row 6: Average 100
                avg_100 = total_200.apply(lambda x: custom_round(x/2))
                for sub in subj_list:
                    block.iloc[6, block.columns.get_loc(sub)] = str(avg_100[sub])

                # Final Summary (Decrease Indent for %)
                grand_total_final = avg_100.sum()
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
            
            st.success("✅ Report Generated Successfully!")
            st.download_button("📥 Download Excel Report", output.getvalue(), "Consolidated_Marksheet.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")
