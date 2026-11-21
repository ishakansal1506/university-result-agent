import os
import streamlit as st
from google import genai
import pandas as pd
import openpyxl

st.set_page_config(page_title="University Result Analysis AI Agent", page_icon="📊", layout="centered")
st.title("📊 University Result Analysis AI Agent")
st.write("Apni Marks file aur University Template upload karein, details bharein, aur ek click mein Excel + Graph download karein.")

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

col1, col2 = st.columns(2)
with col1:
    marks_uploaded = st.file_uploader("1. Upload Student Marks File (.xlsx / .csv)", type=["xlsx", "xls", "csv"])
with col2:
    template_uploaded = st.file_uploader("2. Upload University Template (.xlsx)", type=["xlsx"])

c1, c2 = st.columns(2)
with c1:
    exam_type = st.text_input("Exam Type", placeholder="e.g. CE-1, CE-2, ST-1")
    exam_date = st.text_input("Date of Exam", placeholder="e.g. 20.07.2024")
    sub_code = st.text_input("Subject Code", placeholder="e.g. CS-288")
    faculty_name = st.text_input("Faculty Name", placeholder="e.g. Dr. Gagandeep")
with c2:
    batch_sem = st.text_input("Batch & Semester", placeholder="e.g. CSE-7th")
    max_marks = st.number_input("Maximum Marks", min_value=1.0, value=30.0)
    sub_name = st.text_input("Subject Full Name", placeholder="e.g. Advanced Web Design Frameworks & Development")
    coc_name = st.text_input("CoC Name", value="Dr. Isha Kansal")

if st.button("🚀 Generate Result Analysis Excel", type="primary"):
    if not marks_uploaded or not template_uploaded:
        st.error("Please dono files (Marks file aur University Template) upload karein!")
    else:
        with st.spinner("AI Agent Result Analysis aur Graph bana raha hai..."):
            marks_file = marks_uploaded.name
            template_file = template_uploaded.name
            with open(marks_file, "wb") as f:
                f.write(marks_uploaded.getbuffer())
            with open(template_file, "wb") as f:
                f.write(template_uploaded.getbuffer())

            df_marks = pd.read_excel(marks_file) if marks_file.endswith(('.xlsx', '.xls')) else pd.read_csv(marks_file)
            marks_cols = list(df_marks.columns)
            marks_sample = df_marks.head(10).to_string()
            output_path = "Final_University_Result_Analysis.xlsx"

            prompt = f"""
            You are an expert University Result Analysis Python Automation Agent. Write ONLY valid, runnable Python code (no markdown formatting, no explanations) using pandas and openpyxl.

            Files provided:
            1. Marks File: '{marks_file}'
               Columns: {marks_cols}
               Sample Rows:
               {marks_sample}
            2. University Excel Template File: '{template_file}'

            User Runtime Inputs:
            - Exam Type: '{exam_type}'
            - Batch & Semester: '{batch_sem}'
            - Subject Code: '{sub_code}'
            - Subject Name: '{sub_name}'
            - Date of Exam: '{exam_date}'
            - Maximum Marks: {float(max_marks)}
            - Faculty Name: '{faculty_name}'
            - CoC Name: '{coc_name}'

            Your Exact Tasks:
            1. Read '{marks_file}' and automatically detect the 'Roll Number' column and the 'Obtained Marks' column. Ignore any non-student summary/total rows at the bottom.
            2. Identify ABSENT students FIRST: Any student whose marks are 0, '0', 'AB', 'A', 'Absent' (case-insensitive), or blank/NaN MUST be treated as ABSENT.
               - Collect the Roll Numbers of all such absent students (convert cleanly to string without '.0' decimals) and join them with ', '.
               - Count total absent students.
            3. For PRESENT students (marks > 0), calculate percentage automatically using:
               pct = (float(marks) / {float(max_marks)}) * 100
               Then count how many present students fall into these 7 percentage ranges:
               - Row 7 ('0-39'): 0 < pct < 40
               - Row 8 ('40-49'): 40 <= pct < 50
               - Row 9 ('50-59'): 50 <= pct < 60
               - Row 10 ('60-69'): 60 <= pct < 70
               - Row 11 ('70-79'): 70 <= pct < 80
               - Row 12 ('80-89'): 80 <= pct < 90
               - Row 13 ('90-100'): 90 <= pct <= 100
            4. Load '{template_file}' using openpyxl (preserve all existing formatting, logo, borders, chart, and remaining footer signatures):
               - Update Header Title in cell 'B1' to: f"{exam_type} ({batch_sem}-{sub_code}({sub_name}))"
               - Update Date of Exam in cell 'K3' (or the cell right next to 'Date of Exam') to: '{exam_date}'
               - Write the 7 range counts into cells 'C7', 'C8', 'C9', 'C10', 'C11', 'C12', 'C13'.
               - Ensure cell 'C14' has formula '=SUM(C7:C13)'.
               - Write the total count of Absent students into cell 'C16'.
               - EXACT CELL FOR ABSENT ROLL NUMBERS: Clear cell 'C21' (set value to None) so nothing appears above the box. Write the comma-separated Absentee Roll Numbers string into cell 'C22' (the merged box right next to 'ABSENTEES ROLL NUMBERES' in Row 22) with wrap_text=True.
               - Update Faculty Name in cell 'C25' (the cell right below 'Faculty' in C24) to: '{faculty_name}'
               - Update CoC Name in cell 'E25' (the cell right below 'CoC' in E24) to: '{coc_name}'
               - If the active worksheet has an existing chart (ws._charts), update its title to f"{exam_type}({batch_sem}-{sub_code.split('-')[-1] if '-' in sub_code else sub_code})".
            5. Save the updated workbook as '{output_path}'.
            """

            response = None
            last_error = None
            for model_name in ["gemini-3-flash-preview", "gemini-3.5-flash", "gemini-3.8-flash", "gemini-2.5-flash"]:
                try:
                    response = client.models.generate_content(
                        model=model_name,
                        contents=prompt
                    )
                    break
                except Exception as e:
                    last_error = e
                    continue

            if response is None:
                raise last_error

            clean_code = response.text.replace("```python", "").replace("```", "").strip()
            exec(clean_code)

            with open(output_path, "rb") as file:
                st.success("✅ Result Analysis taiyar hai!")
                st.download_button(
                    label="📥 Download Final Result Analysis (.xlsx)",
                    data=file,
                    file_name="Final_University_Result_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
