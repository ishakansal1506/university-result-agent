import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment

st.set_page_config(page_title="University Result Analysis Generator", page_icon="📊", layout="centered")
st.title("📊 University Result Analysis Generator")
st.write("Apni Marks file aur University Template upload karein, details bharein, aur 1 second mein Excel + Graph download karein.")

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

def find_roll_and_marks_cols(df):
    cols = list(df.columns)
    roll_col = None
    marks_col = None
    for col in cols:
        c_low = str(col).strip().lower()
        if any(k in c_low for k in ("roll", "univ", "id", "reg")):
            roll_col = col
        elif any(k in c_low for k in ("mark", "score", "obt", "total", "ce", "st", "pct", "%")):
            marks_col = col
    if roll_col is None:
        roll_col = next(iter(cols))
    if marks_col is None:
        marks_col = cols.pop() if len(cols) > 1 else roll_col
    return roll_col, marks_col

if st.button("🚀 Generate Result Analysis Excel", type="primary"):
    if not marks_uploaded or not template_uploaded:
        st.error("Please dono files (Marks file aur University Template) upload karein!")
    else:
        try:
            marks_file = marks_uploaded.name
            template_file = template_uploaded.name
            with open(marks_file, "wb") as f:
                f.write(marks_uploaded.getbuffer())
            with open(template_file, "wb") as f:
                f.write(template_uploaded.getbuffer())

            df = pd.read_excel(marks_file) if marks_file.endswith(('.xlsx', '.xls')) else pd.read_csv(marks_file)
            df = df.dropna(how="all")
            roll_col, marks_col = find_roll_and_marks_cols(df)

            absent_rolls = list()
            r0_39 = 0
            r40_49 = 0
            r50_59 = 0
            r60_69 = 0
            r70_79 = 0
            r80_89 = 0
            r90_100 = 0

            for _, row in df.iterrows():
                raw_roll = row.get(roll_col)
                raw_mark = row.get(marks_col)

                if pd.isna(raw_roll):
                    continue
                roll_str = str(raw_roll).strip()
                if roll_str.endswith(".0"):
                    roll_str = roll_str[:-2]
                if roll_str.lower() in ("total", "grand total", "average", "avg", "roll no", "roll number", "s.no", "sr.no"):
                    continue

                is_absent = False
                mark_val = 0.0
                if pd.isna(raw_mark):
                    is_absent = True
                else:
                    m_str = str(raw_mark).strip().lower()
                    if m_str in ("", "ab", "a", "absent", "nan", "none", "-", "0", "0.0"):
                        is_absent = True
                    else:
                        try:
                            mark_val = float(m_str)
                            if mark_val <= 0:
                                is_absent = True
                        except ValueError:
                            is_absent = True

                if is_absent:
                    absent_rolls.append(roll_str)
                else:
                    pct = (mark_val / float(max_marks)) * 100.0
                    if pct < 40:
                        r0_39 += 1
                    elif pct < 50:
                        r40_49 += 1
                    elif pct < 60:
                        r50_59 += 1
                    elif pct < 70:
                        r60_69 += 1
                    elif pct < 80:
                        r70_79 += 1
                    elif pct < 90:
                        r80_89 += 1
                    else:
                        r90_100 += 1

            wb = openpyxl.load_workbook(template_file)
            ws = wb.active

            ws["B1"].value = f"{exam_type} ({batch_sem}-{sub_code}({sub_name}))"
            ws["K3"].value = exam_date

            ws["C7"].value = r0_39
            ws["C8"].value = r40_49
            ws["C9"].value = r50_59
            ws["C10"].value = r60_69
            ws["C11"].value = r70_79
            ws["C12"].value = r80_89
            ws["C13"].value = r90_100
            ws["C14"].value = "=SUM(C7:C13)"

            ws["C16"].value = len(absent_rolls)
            ws["C21"].value = None
            ws["C22"].value = ", ".join(absent_rolls) if absent_rolls else "Nil"
            ws["C22"].alignment = Alignment(wrap_text=True, vertical="center")

            ws["C25"].value = faculty_name
            ws["E25"].value = coc_name

            short_code = sub_code.split("-").pop() if "-" in sub_code else sub_code
            for chart in ws._charts:
                chart.title = f"{exam_type}({batch_sem}-{short_code})"

            output_path = "Final_University_Result_Analysis.xlsx"
            wb.save(output_path)

            with open(output_path, "rb") as file:
                total_present = r0_39 + r40_49 + r50_59 + r60_69 + r70_79 + r80_89 + r90_100
                st.success(f"✅ 1 Second mein taiyar! Present: {total_present} | Absent/Zero: {len(absent_rolls)}")
                st.download_button(
                    label="📥 Download Final Result Analysis (.xlsx)",
                    data=file,
                    file_name="Final_University_Result_Analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as err:
            st.error(f"Error: {err}")
