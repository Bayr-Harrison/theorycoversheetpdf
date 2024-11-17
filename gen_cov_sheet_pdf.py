import os
import streamlit as st
import pandas as pd
import pg8000
from io import BytesIO
import zipfile
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

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

# Function to generate a protected and formatted Excel sheet
def create_protected_excel_sheet(student_data, student_id):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Coversheet"

    # Styling variables
    header_font = Font(bold=True)
    cell_alignment = Alignment(horizontal="center", vertical="center")
    thin_border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
    header_fill = PatternFill(start_color="AEE1F8", end_color="AEE1F8", fill_type="solid")

    # Populate static text in specific cells
    sheet["B2"] = "Student Name:"
    sheet["B3"] = "Student IATC ID:"
    sheet["B4"] = "Student National ID:"
    sheet["B5"] = "Student Class:"

    # Apply formatting to static cells
    for cell in ["B2", "B3", "B4", "B5"]:
        sheet[cell].font = header_font
        sheet[cell].alignment = cell_alignment
        sheet[cell].fill = header_fill
        sheet[cell].border = thin_border

    # Populate student-specific values
    sheet["C2"] = student_data['Name'].iloc[0]
    sheet["C3"] = student_data['IATC ID'].iloc[0]
    sheet["C4"] = student_data['National ID'].iloc[0]
    sheet["C5"] = float(student_data['Class'].iloc[0])  # Convert to float for decimal formatting

    for cell in ["C2", "C3", "C4", "C5"]:
        sheet[cell].alignment = cell_alignment
        sheet[cell].border = thin_border

    # Format C5 to display as a number with 2 decimal places
    sheet["C5"].number_format = "0.00"

    # Populate table headers
    headers = ['Subject', 'Score', 'Result', 'Date']
    for col_num, header in enumerate(headers, start=2):
        cell = sheet.cell(row=7, column=col_num, value=header)
        cell.font = header_font
        cell.alignment = cell_alignment
        cell.border = thin_border
        cell.fill = header_fill

    # Populate the data table
    for row_num, row_data in enumerate(student_data[['Subject', 'Score', 'Result', 'Date']].values, start=8):
        for col_num, value in enumerate(row_data, start=2):
            cell = sheet.cell(row=row_num, column=col_num, value=value)
            cell.alignment = cell_alignment
            cell.border = thin_border

    # Adjust column widths to fit content
    for col in range(2, 6):  # Columns B to E
        max_length = 0
        for row in range(1, sheet.max_row + 1):
            cell_value = sheet.cell(row=row, column=col).value
            if cell_value:
                max_length = max(max_length, len(str(cell_value)))
        adjusted_width = max_length + 2  # Add padding
        sheet.column_dimensions[get_column_letter(col)].width = adjusted_width

    # Protect the sheet
    sheet.protection.set_password(os.environ["EXCEL_PASSWORD"])  # Set a password for protection
    sheet.protection.enable()  # Enable sheet protection

    # Save the workbook to a temporary file
    excel_buffer = BytesIO()
    workbook.save(excel_buffer)
    excel_buffer.seek(0)
    return excel_buffer

# Function to generate Excel files with protection and save them to a zip file
def generate_coversheets_zip(student_list):
    df = fetch_student_data(student_list)

    # Create an in-memory ZIP file
    zip_buffer = BytesIO()
    with zipfile.ZipFile(zip_buffer, "w") as zip_file:
        for student_id in student_list:
            filtered_df = df[df['IATC ID'] == student_id]

            # Generate protected and formatted Excel file
            excel_buffer = create_protected_excel_sheet(filtered_df, student_id)

            # Add Excel file to the zip
            zip_file.writestr(f"{student_id}.xlsx", excel_buffer.read())

    zip_buffer.seek(0)
    return zip_buffer

# Streamlit interface
st.title("Generate Theory Exam Coversheets (Protected & Formatted)")
st.write("Enter a list of student IDs and download the Excel coversheets containing the highest result for each subject the student has taken.")

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
