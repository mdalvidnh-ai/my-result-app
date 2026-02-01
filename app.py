import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Result Editor", layout="wide")

st.title("📝 Exam Mark Entry & Consolidation")
st.write("1. Upload Excel | 2. Enter Practical Marks in the table below | 3. Download")

def custom_round(x):
    return int(np.floor(x + 0.5))

uploaded_file = st.file_uploader("Upload your Exam Workbook", type="xlsx")

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        exam_names = ['FIRST UNIT TEST', 'FIRST TERM', 'SECOND UNIT TEST', 'ANNUAL EXAM']
        subj_cols = {'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'}
        
        all_students = {}

        # 1. Extract Data from Sheets
        for sheet in exam_names:
            if sheet in xl.sheet_names:
                df = xl.parse(sheet)
                df.columns = df.columns.str.strip().str.upper()
                for _, row in df.iterrows():
                    roll = row['ROLL NO.']
                    if roll not in all_students:
                        all_students[roll] = {'Name': row.get('STUDENT NAME', 'Unknown'), 'Exams': {}}
                    all_students[roll]['Exams'][sheet] = {target: row.get(source, 0) for source, target in subj_cols.items()}

        # 2. Build the Initial Multi-Row DataFrame
        rows = []
        for roll in sorted(all_students.keys()):
            s = all_students[roll]
            categories = ['FIRST UNIT TEST (25)', 'FIRST TERM EXAM (50)', 'SECOND UNIT TEST (25)', 'ANNUAL EXAM (70/80)', 'INT/PRACTICAL (20/30)']
            
            for cat in categories:
                row_data = {
                    'Roll No.': roll if cat == 'FIRST UNIT TEST (25)' else '',
                    'Student Name': s['Name'] if cat == 'FIRST UNIT TEST (25)' else '',
                    'Exam Type': cat
                }
                # Pre-fill marks from uploaded sheets
                key = cat.split(' (')[0] # Get 'FIRST TERM' etc.
                if key in s['Exams']:
                    row_data.update(s['Exams'][key])
                else:
                    # Fill 0 for subjects if it's the Practical row
                    for sub in subj_cols.values(): row_data[sub] = 0
                rows.append(row_data)

        base_df = pd.DataFrame(rows)

        # 3. INTERACTIVE EDITING SECTION
        st.subheader("Edit Practical Marks Directly Below:")
        st.info("💡 You can click on any cell in the 'INT/PRACTICAL' rows to type in marks.")
        
        # This makes the table editable
        edited_df = st.data_editor(
            base_df,
            column_config={
                "Roll No.": st.column_config.Column(disabled=True),
                "Student Name": st.column_config.Column(disabled=True),
                "Exam Type": st.column_config.Column(disabled=True),
            },
            disabled=["Roll No.", "Student Name", "Exam Type"],
            hide_index=True,
            use_container_width=True
        )

        # 4. RECALCULATION LOGIC
        if st.button("Calculate Totals & Prepare Download"):
            final_data = []
            # Group by 5 rows (one student block)
            for i in range(0, len(edited_df), 5):
                student_block = edited_df.iloc[i:i+5]
                
                # Add the original 5 rows
                for _, row in student_block.iterrows():
                    final_data.append(row.to_dict())
                
                # Calculate Sum Out of 200
                sum_200 = {'Roll No.': '', 'Student Name': '', 'Exam Type': 'Total Marks Out of 200'}
                avg_100 = {'Roll No.': '', 'Student Name': '', 'Exam Type': 'Average Marks 200/2=100'}
                
                for sub in subj_cols.values():
                    total_val = pd.to_numeric(student_block[sub]).sum()
                    sum_200[sub] = total_val
                    avg_100[sub] = custom_round(total_val / 2)
                
                final_data.append(sum_200)
                final_data.append(avg_100)

            # Final Output DataFrame
            final_df = pd.DataFrame(final_data)
            
            # 5. EXPORT
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Consolidated')
            
            st.success("Calculations Finished!")
            st.download_button(
                label="📥 Download Final Consolidated Sheet",
                data=output.getvalue(),
                file_name="Final_Consolidated_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Error: {e}")
