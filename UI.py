import streamlit as st
from io import BytesIO
import hashlib
import process_attendance  # <- import your script

st.title("Excel Attendance Processor")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls", "xlsm"])

if uploaded_file:
    # Compute hash of uploaded file to detect changes
    uploaded_file.seek(0)  # Ensure at start
    file_content = uploaded_file.read()
    file_hash = hashlib.md5(file_content).hexdigest()
    uploaded_file.seek(0)  # Reset file pointer after hashing

    if (
        "file_hash" not in st.session_state
        or st.session_state.file_hash != file_hash
        or "processed_output" not in st.session_state
    ):
        try:
            # Initialize progress bar and status text
            progress_bar = st.progress(0)
            status_text = st.empty()

            # Define a callback function to update progress
            def update_progress(progress):
                progress_bar.progress(progress)
                if progress == 0:
                    status_text.text("Starting processing...")
                elif progress == 20:
                    status_text.text("Deleting absent rows...")
                elif progress == 80:
                    status_text.text("Filling missing values and finalizing...")
                elif progress == 100:
                    status_text.text("Preparing download...")

            # Process the Excel file with progress callback
            wb = process_attendance.process_excel(
                uploaded_file, progress_callback=update_progress
            )

            # Save to BytesIO
            output = BytesIO()
            wb.save(output)
            output.seek(0)

            # Store in session state
            st.session_state.processed_output = output.getvalue()
            st.session_state.file_hash = file_hash

            progress_bar.progress(100)
            status_text.text("Processing complete!")
            st.success("File processed successfully!")

            # Clear status text
            status_text.empty()

        except Exception as e:
            st.error(f"Error processing file: {e}")
            if "progress_bar" in locals():
                progress_bar.empty()
            if "status_text" in locals():
                status_text.empty()

    # Show download button using cached output
    if "processed_output" in st.session_state:
        st.download_button(
            label="Download Processed Excel",
            data=st.session_state.processed_output,
            file_name=f"{uploaded_file.name.split('.')[0]}_processed.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
