import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="School Result Consolidator", layout="wide")

st.title("🏫 Student Result Consolidation Tool")
st.write("Upload your Excel file. This app will populate the 'Consolidated' sheet automatically.")

uploaded_file = st.file_uploader("Choose your Excel file", type="xlsx")

if uploaded_file:
    try:
        # 1. Load the data
        xl = pd.ExcelFile(uploaded_file)
        
        # Mapping subject names from Exam Sheets to Consolidated Sheet
        subj_map = {
            'ENG': 'Eng', 'MAR': 'Mar', 'GEOG': 'Geo', 
            'SOCIO': 'Soc', 'PSYC': 'Psy', 'ECO': 'Eco'
        }
        exams = {
            'FIRST UNIT TEST': 'FIRST UNIT TEST (25)',
            'FIRST TERM': 'FIRST TERM EXAM (50)',
            'SECOND UNIT TEST': 'SECOND UNIT TEST (25)',
            'ANNUAL EXAM': 'ANNUAL EXAM (70/80)'
        }

        # 2. Extract Data from all sheets
        all_data = {}
        for sheet_name in exams.keys():
            if sheet_name in xl.sheet_names:
                df = xl.parse(sheet_name)
                df.columns = df.columns.str.strip().str.upper()
                # Store data by Roll No
                for _, row in df.iterrows():
                    roll = row['ROLL NO.']
                    if roll not in all_data:
                        all_data[roll] = {'Name': row.get('STUDENT NAME', 'Unknown')}
                    
                    all_data[roll][sheet_name] = {
                        target: row.get(source, 0) for source, target in subj_map.items()
                    }

        # 3. Create the Consolidated Structure
        rows = []
        sorted_rolls = sorted(all_data.keys())
        
        for roll in sorted_rolls:
            student = all_data[roll]
            # Row 1: First Unit Test
            r1 = {'Roll No.': roll, 'Column1': student['Name'], 'Column2': 'FIRST UNIT TEST (25)'}
            r1.update(student.get('FIRST UNIT TEST', {}))
            rows.append(r1)
            
            # Row 2: First Term
            r2 = {'Roll No.': '', 'Column1': '', 'Column2': 'FIRST TERM EXAM (50)'}
            r2.update(student.get('FIRST TERM', {}))
            rows.append(r2)
            
            # Row 3: Second Unit Test
            r3 = {'Roll No.': '', 'Column1': '', 'Column2': 'SECOND UNIT TEST (25)'}
            r3.update(student.get('SECOND UNIT TEST', {}))
            rows.append(r3)
            
            # Row 4: Annual Exam
            r4 = {'Roll No.': '', 'Column1': '', 'Column2': 'ANNUAL EXAM (70/80)'}
            r4.update(student.get('ANNUAL EXAM', {}))
            rows.append(r4)
            
            # Add rows for Practical/Totals (Placeholders as per your template)
            rows.append({'Roll No.': '', 'Column1': '', 'Column2': 'INT/PRACTICAL (20/30)'})
            rows.append({'Roll No.': '', 'Column1': '', 'Column2': 'Total Marks Out of 200'})
            rows.append({'Roll No.': '', 'Column1': '', 'Column2': 'Average Marks 200/2=100'})

        consolidated_df = pd.DataFrame(rows)

        # 4. Show Preview and Download
        st.success("✅ Consolidation Complete!")
        st.dataframe(consolidated_df.head(14)) # Show first 2 students

        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            consolidated_df.to_excel(writer, index=False, sheet_name='Consolidated')
        
        st.download_button(
            label="📥 Download Consolidated Excel",
            data=output.getvalue(),
            file_name="Final_Consolidated_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}")