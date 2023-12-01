import os
import re
import win32com.client
from docx2pdf import convert
from fpdf import FPDF
import PyPDF2

result_path = 'Result'
secret_path = 'Secret'

def encrypt_pdf():
    for file in os.listdir(result_path):
        if file.endswith(".pdf"):
            input_pdf = os.path.join(result_path, file)

            # Enter PDF file password
            pdf = PyPDF2.PdfReader(open(input_pdf, "rb"))
            password = input(f"Enter the password for '{file}': ")

            # Encrypt the PDF
            encrypted_pdf = PyPDF2.PdfWriter()
            encrypted_pdf.append_pages_from_reader(pdf)
            encrypted_pdf.encrypt(password)

            # Save the encrypted PDF in the "Secret" folder with the file name + "_encrypted"
            encrypted_file = os.path.splitext(file)[0] + "_encrypted.pdf"
            output_pdf = os.path.join(secret_path, encrypted_file)

            with open(output_pdf, "wb") as output_file:
                encrypted_pdf.write(output_file)

            print(f"Encrypted file saved: {secret_path}")


encrypt_pdf()