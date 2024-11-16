import os
import streamlit as st
import pandas as pd
import pg8000
from io import BytesIO
import zipfile
from fpdf import FPDF

# Function to generate coversheets and save them as PDFs in a zip file
def generate_coversheets_zip(student_list=[]):
    # Connect to the database
    db_connection = pg8000.connect(
        database=os.environ["SUPABASE_DB_NAME"],
        user=os.environ["SUPABASE_USER"],
        password=os.environ["SUPABASE_PASSWORD"],
        host=os.environ["SUPABASE_HOST"],
        port=os.environ["SUPABASE_PORT"]
    )
    db_cursor = db_connection.cursor()

    # Prepare SQL query
    student_list_string = ', '.join(map(str, student_list))
    db_query = f"""
    SELECT student_list.name,                  
           student_list.iatc_id,
           student_list.nat_id,
           student_list.class,
           exam_list.exam_long AS subject,
           exam_results.score,
           exam_results.result,
           exam_results.date
    FROM exam_results 
    JOIN student_list ON exam_results.nat_id = student_list.nat_id
    JOIN exam_list ON exam_results.exam = exam_list.exam
    WHERE student_list.iatc_id IN ({student_list_string}) 
          AND exam_results.score_index = 1
    ORDER BY exam_list.srt_exam ASC
    """
    db_cursor.execute(db_query)
    output_data = db_cursor.fetchall()
    db_cursor.close()
    db_connection.close()

    # Create DataFrame
    col_names = ['Name', 'IATC ID', 'National ID', 'Class', 'Subject', 'Score', 'Result', 'Date']
    df = pd.DataFrame(output_data, columns=col_names)

    # Create an in-memory ZIP file
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for student_id in student_list:
            filtered_df = df[df['IATC ID'] == student_id]

            # Create a new PDF for each student
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            # Add student details
            pdf.cell(200, 10, txt="Student Coversheet", ln=True, align="C")
            pdf.ln(10)
            pdf.cell(50, 10, txt="Student Name:", border=1)
            pdf.cell(100, 10, txt=str(filtered_df['Name'].iloc[0]), border=1, ln=True)
            pdf.cell(50, 10, txt="IATC ID:", border=1)
            pdf.cell(100, 10, txt=str(filtered_df['IATC ID'].iloc[0]), border=1, ln=True)
            pdf.cell(50, 10, txt="National ID:", border=1)
            pdf.cell(100, 10, txt=str(filtered_df['National ID'].iloc[0]), border=1, ln=True)
            pdf.cell(50, 10, txt="Class:", border=1)
            pdf.cell(100, 10, txt=str(filtered_df['Class'].iloc[0]), border=1, ln=True)

            # Add table headers
            pdf.ln(10)
            for header in ['Subject', 'Score', 'Result', 'Date']:
                pdf.cell(48, 10, txt=header, border=1, align="C")
            pdf.ln()

            # Add table rows
            for _, row in filtered_df.iterrows():
                pdf.cell(48, 10, txt=str(row['Subject']), border=1)
                pdf.cell(48, 10, txt=str(row['Score']), border=1)
                pdf.cell(48, 10, txt=str(row['Result']), border=1)
                pdf.cell(48, 10, txt=str(row['Date']), border=1)
                pdf.ln()

            # Save PDF to BytesIO buffer
            pdf_buffer = BytesIO()
            pdf.output(pdf_buffer)  # Output PDF content to the buffer
            pdf_buffer.seek(0)

            # Write PDF to the zip file
            zip_file.writestr(f"{student_id}.pdf", pdf_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer

# Streamlit interface
st.title("Generate Theory Exam Coversheets")
st.write("Enter a list of student IDs and download the coversheets as PDFs.")

student_ids_input = st.text_area("Enter Student IDs separated by commas (e.g., 151596, 156756, 154960):")
st.write("Need help generating a list of IDs? Download the Excel template:")

# Direct link to the Excel file in GitHub
template_url = "https://github.com/Bayr-Harrison/coversheetgenerator/raw/main/Coversheet%20Generator%20Input.xlsx"
st.markdown(f"[Download Excel Template]({template_url})", unsafe_allow_html=True)

if st.button("Generate Coversheets"):
    try:
        student_list = [int(id.strip()) for id in student_ids_input.split(",")]
        st.write("Generating coversheets...")

        # Generate the zip file in memory
        zip_file = generate_coversheets_zip(student_list)

        # Offer the zip file for download
        st.download_button(
            label="Download All Coversheets as ZIP",
            data=zip_file,
            file_name="coversheets.zip",
            mime="application/zip"
        )
        st.success("Coversheets zip generated successfully!")

    except Exception as e:
        st.error(f"An error occurred: {e}")
