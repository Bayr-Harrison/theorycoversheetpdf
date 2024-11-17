import os
import streamlit as st
import pandas as pd
import pg8000
from io import BytesIO
import zipfile
from fpdf import FPDF

# Function to establish a database connection
def get_database_connection():
    db_connection = pg8000.connect(
        database=os.environ["SUPABASE_DB_NAME"],
        user=os.environ["SUPABASE_USER"],
        password=os.environ["SUPABASE_PASSWORD"],
        host=os.environ["SUPABASE_HOST"],
        port=os.environ["SUPABASE_PORT"]
    )
    return db_connection

# Function to query the database and fetch data
def fetch_student_data(student_list):
    db_connection = get_database_connection()
    db_cursor = db_connection.cursor()
    
    student_list_string = ', '.join(map(str, student_list))

    db_query = f"""SELECT student_list.name,                  
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
                    WHERE student_list.iatc_id IN ({student_list_string}) AND exam_results.score_index = 1
                    ORDER BY exam_list.srt_exam ASC
                """
    db_cursor.execute(db_query)
    output_data = db_cursor.fetchall()
    db_cursor.close()
    db_connection.close()

    col_names = ['Name', 'IATC ID', 'National ID', 'Class', 'Subject', 'Score', 'Result', 'Date']
    df = pd.DataFrame(output_data, columns=col_names)
    return df

# Function to generate a PDF for a student
def generate_pdf(student_data, student_id):
    pdf = FPDF()
    pdf.add_page()

    # Set default font and add title
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, f"Coversheet for {student_data['Name'].iloc[0]}", align='C', ln=True)

    # Add student details
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, f"Student Name: {student_data['Name'].iloc[0]}", ln=True)
    pdf.cell(0, 10, f"IATC ID: {student_data['IATC ID'].iloc[0]}", ln=True)
    pdf.cell(0, 10, f"National ID: {student_data['National ID'].iloc[0]}", ln=True)
    pdf.cell(0, 10, f"Class: {student_data['Class'].iloc[0]}", ln=True)

    pdf.ln(10)  # Line break

    # Add table header
    pdf.set_font('Arial', 'B', 10)
    headers = ['Subject', 'Score', 'Result', 'Date']
    for header in headers:
        pdf.cell(40, 10, header, border=1, align='C')
    pdf.ln()

    # Add table rows
    pdf.set_font('Arial', '', 10)
    for _, row in student_data[['Subject', 'Score', 'Result', 'Date']].iterrows():
        pdf.cell(40, 10, str(row['Subject']), border=1)
        pdf.cell(40, 10, str(row['Score']), border=1)
        pdf.cell(40, 10, str(row['Result']), border=1)
        pdf.cell(40, 10, str(row['Date']), border=1)
        pdf.ln()

    # Save to buffer
    pdf_buffer = BytesIO()
    pdf.output(pdf_buffer)
    pdf_buffer.seek(0)

    return pdf_buffer

# Function to generate PDFs and save them to a zip file
def generate_coversheets_zip(student_list):
    df = fetch_student_data(student_list)

    # Create an in-memory ZIP file
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for student_id in student_list:
            filtered_df = df[df['IATC ID'] == student_id]

            # Generate PDF for the student
            pdf_buffer = generate_pdf(filtered_df, student_id)

            # Add PDF to the zip
            pdf_filename = f"{student_id}.pdf"
            zip_file.writestr(pdf_filename, pdf_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer

# Streamlit interface
st.title("Generate Theory Exam Coversheets (PDF)")
st.write("Enter a list of student IDs and download the PDF coversheets containing the highest result for each subject the student has taken.")

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
