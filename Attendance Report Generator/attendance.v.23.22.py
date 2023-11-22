import time
import pandas as pd
from fpdf import FPDF

class PDF(FPDF):
    def footer(self):
        self.set_y(-15)
        self.set_font('Helvetica', 'I', 8)
        self.set_text_color(0,0,0)
        self.cell(0, 10, "Created By Vshker M", 0, 0, 'C') #'Page ' + str(self.page_no()) + 

def get_user_id():
    # Get user input for ID
    user_id = input("Enter your Student ID: ")

    # Validation for the user ID
    while not user_id.isdigit():
        print("Invalid input. Please enter a numeric user ID.")
        user_id = input("Enter your Student ID:")

    # Convert the user ID to an integer
    return int(user_id)

def create_letterhead(pdf, width):
    pdf.image("/home/bookertee/Desktop/Abhayaz_drive/python_project/python_automation/images/Abhyaz.logo.jpg", 150, 10, 50)

def create_title(title, pdf):
    pdf.set_font('Helvetica', 'b', 20)
    pdf.ln(40)
    pdf.write(5, title)
    pdf.ln(10)

    # Add date of report
    pdf.set_font('Helvetica', '', 10)
    pdf.set_text_color(r=0, g=0, b=0)
    today = time.strftime("%d/%m/%Y")
    pdf.write(4, f'{today}')

    # Add line break
    pdf.ln(10)

def create_pdf(user_id, excel_file):
    try:
        # Load Excel data into a pandas DataFrame, starting from the 4th row
        df = pd.read_excel(excel_file, header=3)

        # Select relevant information based on user_id
        user_info = df[df['Student ID'] == user_id][['First name', 'Last name', 'Email address', 'Student ID', 'P', 'L', 'E', 'A', 'Taken sessions', 'Points', 'Percentage']]

        if user_info.empty:
            print("User not found")
        else:
           # Transpose the dataframe to have headings in columns
            user_info = user_info.transpose()

            # Create a PDF document
            pdf_file = f'result/student_info_{user_id}.pdf'
            pdf = PDF()
            pdf.add_page()

            # Add letterhead and title
            create_letterhead(pdf, 50)
            create_title('ATTENDANCE REPORT', pdf)

            # Add user information to the PDF in tabular format
            pdf.ln(10)
            pdf.set_font("Helvetica", "B", size=12)
            pdf.cell(0, 10, f"Attendances for Student ID {user_id}:", ln=True)

            # Add table header in columns
            pdf.ln(10)
            pdf.set_font("Helvetica", "B", size=10)
            for col_name, col_value in user_info.iterrows():
                # Add a cell for col_name with borders
                pdf.cell(80, 10, f"{col_name}", border=True, align= "C")

                # Add a cell for col_value with borders
                pdf.cell(80, 10, f"{col_value.values[0]}", ln=True, border=True, align= "C")
            pdf.ln()

            # Add some words to PDF
            pdf.ln(10)
            pdf.set_font("Helvetica", size=12)
            pdf.cell(0, 10, "Note:", ln=True)
            pdf.multi_cell(50, 6, "P:Present \nL:LATE \nE:EXCUSE \nA:ABSENT",border=True)

            # Generate the PDF
            pdf.output(pdf_file, 'F')
            print(f'PDF created: {pdf_file}')

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    excel_file_path = '/home/bookertee/Desktop/Abhayaz_drive/python_project/python_automation/data/Engineering Intern Submission_Attendances_20231110-1759.xlsx'
    create_pdf(get_user_id(), excel_file_path)
