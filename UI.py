import streamlit as st
from io import BytesIO
import process_attendance  # <- import your script

st.title("Excel Attendance Processor")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls", "xlsm"])

if uploaded_file:
    try:
        wb = process_attendance.process_excel(uploaded_file)
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        st.success("File processed successfully!")
        st.download_button(
            label="Download Processed Excel",
            data=output,
            file_name=f"{uploaded_file.name.split('.')[0]}_processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"Error processing file: {e}")
