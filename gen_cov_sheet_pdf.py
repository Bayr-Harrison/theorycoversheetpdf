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

    # Create an in-memory ZIP file for PDFs
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for student_id in student_list:
            filtered_df = df[df['IATC ID'] == student_id]

            # Create PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)

            # Add student details
            pdf.set_fill_color(174, 225, 248)  # Light blue fill color
            pdf.set_text_color(0, 0, 0)       # Black text color
            pdf.set_font(style="B")           # Bold font

            # Header fields
            pdf.cell(50, 10, "Student Name:", fill=True, border=1)
            pdf.cell(100, 10, filtered_df['Name'].iloc[0], border=1, ln=1)

            pdf.cell(50, 10, "Student IATC ID:", fill=True, border=1)
            pdf.cell(100, 10, str(filtered_df['IATC ID'].iloc[0]), border=1, ln=1)

            pdf.cell(50, 10, "Student National ID:", fill=True, border=1)
            pdf.cell(100, 10, str(filtered_df['National ID'].iloc[0]), border=1, ln=1)

            pdf.cell(50, 10, "Student Class:", fill=True, border=1)
            pdf.cell(100, 10, filtered_df['Class'].iloc[0], border=1, ln=1)

            pdf.ln(10)  # Add spacing

            # Add table headers
            headers = ['Subject', 'Score', 'Result', 'Date']
            pdf.set_font(style="B")
            pdf.set_fill_color(174, 225, 248)  # Light blue for headers
            for header in headers:
                pdf.cell(45, 10, header, border=1, fill=True, align="C")
            pdf.ln()

            # Add table rows
            pdf.set_font(style="")
            for row in filtered_df[['Subject', 'Score', 'Result', 'Date']].values:
                for item in row:
                    pdf.cell(45, 10, str(item), border=1, align="C")
                pdf.ln()

            # Save the PDF to a buffer
            pdf_buffer = BytesIO()
            pdf.output(pdf_buffer)
            pdf_buffer.seek(0)

            # Add the PDF to the ZIP file
            pdf_filename = f"{student_id}.pdf"
            zip_file.writestr(pdf_filename, pdf_buffer.read())

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
