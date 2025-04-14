import pandas as pd
from fpdf import FPDF
import yagmail
import os
from dotenv import load_dotenv

# Load email and password from .env file
load_dotenv()
EMAIL = os.getenv("kelvinanesumutukwa18@gmail.com")
PASSWORD = os.getenv("olmo rspd udnp udjc")

# Read employee data from Excel
try:
    df = pd.read_excel("employees.xlsx")
    df.columns = df.columns.str.strip()
 # Remove any leading or trailing spaces
except FileNotFoundError:
    print("⚠ Error: 'employees.xlsx' not found.")
    exit()

# Make sure output folder exists
os.makedirs("payslips", exist_ok=True)

# Loop through each employee
for index, row in df.iterrows():
    emp_id = row['Employee ID']
    name = row['Name']
    email = row['Email']
    basic = row['Basic Salary']
    allowance = row['Allowances']
    deduction = row['Deductions']
    net = basic + allowance - deduction

    # Create PDF payslip
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)

    pdf.cell(200, 10, txt="Payslip", ln=True, align='C')
    pdf.cell(200, 10, txt=f"Employee Name: {name}", ln=True)
    pdf.cell(200, 10, txt=f"Employee ID: {emp_id}", ln=True)
    pdf.cell(200, 10, txt=f"Basic Salary: ${basic}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: ${allowance}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: ${deduction}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: ${net}", ln=True)

    # Save PDF file
    pdf_path = f"payslips/{emp_id}.pdf"
    pdf.output(pdf_path)

    # Send email with payslip
    try:
        yag = yagmail.SMTP("kelvinanesumutukwa18@gmail.com", 'olmo rspd udnp udjc')
        yag.send(
            to=email,
            subject="Your Payslip for This Month",
            contents=f"Hi {name},\n\nPlease find your payslip attached.\n\nBest regards,\nHR Team",
            attachments=pdf_path
        )
        print(f"✅ Email sent to {name} ({email})")
    except Exception as e:
        print(f"❌ Failed to send email to {email}: {e}")