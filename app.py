import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Result Processor", layout="wide")

st.title("📊 School Exam Consolidator")
st.write("Upload your Excel workbook to generate the automatic summary.")

uploaded_file = st.file_uploader("Upload App.xlsx", type="xlsx")

def custom_round(x):
    # Round up if decimal is >= 0.5 (e.g., 1.5 -> 2.0)
    return np.floor(x + 0.5)

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        exam_names = ['FIRST UNIT TEST', 'FIRST TERM', 'SECOND UNIT TEST', 'ANNUAL EXAM']
        
        # Mapping subjects from your sheets to the template
        subj_cols = {'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'}
        extra_cols = {'TOTAL': 'Grand Total', '%': '%', 'RESULT': 'Result'}

        all_students = {}

        # 1. Collect Data
        for sheet in exam_names:
            if sheet in xl.sheet_names:
                df = xl.parse(sheet)
                df.columns = df.columns.str.strip().str.upper()
                
                for _, row in df.iterrows():
                    roll = row['ROLL NO.']
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    
                    # Store marks and existing totals
                    marks = {target: row.get(source, 0) for source, target in subj_cols.items()}
                    stats = {target: row.get(source, 0) for source, target in extra_cols.items()}
                    all_students[roll]['Exams'][sheet] = {**marks, **stats}

        # 2. Build the Multi-Row Template
        final_rows = []
        for roll in sorted(all_students.keys()):
            s = all_students[roll]
            
            # Row labels in the order they appear in your template
            categories = [
                ('FIRST UNIT TEST (25)', 'FIRST UNIT TEST'),
                ('FIRST TERM EXAM (50)', 'FIRST TERM'),
                ('SECOND UNIT TEST (25)', 'SECOND UNIT TEST'),
                ('ANNUAL EXAM (70/80)', 'ANNUAL EXAM'),
                ('INT/PRACTICAL (20/30)', None) # Placeholder for manual entry
            ]

            student_marks_sum = {sub: 0 for sub in subj_cols.values()}

            for label, key in categories:
                row_data = {
                    'Roll No.': roll if label == 'FIRST UNIT TEST (25)' else '',
                    'Column1': s['Name'] if label == 'FIRST UNIT TEST (25)' else '',
                    'Column2': label
                }
                
                if key and key in s['Exams']:
                    exam_data = s['Exams'][key]
                    row_data.update(exam_data)
                    # Add to running sum for the "Out of 200" row
                    for sub in student_marks_sum:
                        try:
                            student_marks_sum[sub] += float(exam_data.get(sub, 0))
                        except: pass
                
                final_rows.append(row_data)

            # ADD CALCULATION ROWS
            # Row: Total Marks Out of 200
            total_200_row = {'Roll No.': '', 'Column1': '', 'Column2': 'Total Marks Out of 200'}
            total_200_row.update(student_marks_sum)
            final_rows.append(total_200_row)

            # Row: Average Marks 200/2=100 (With Custom Rounding)
            avg_100_row = {'Roll No.': '', 'Column1': '', 'Column2': 'Average Marks 200/2=100'}
            for sub, val in student_marks_sum.items():
                avg_100_row[sub] = custom_round(val / 2)
            final_rows.append(avg_100_row)

        # 3. Create DataFrame
        output_df = pd.DataFrame(final_rows)
        
        st.success("Successfully Processed!")
        st.dataframe(output_df)

        # Download Button
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            output_df.to_excel(writer, index=False, sheet_name='Consolidated')
        
        st.download_button("📥 Download Final Sheet", output.getvalue(), "Consolidated_Report.xlsx")

    except Exception as e:
        st.error(f"Error: {e}. Make sure column names like 'ROLL NO.' are correct.")
