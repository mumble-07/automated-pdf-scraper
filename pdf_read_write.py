# -*- coding: utf-8 -*-
"""
Created on Tues Feb 06 9:38:29 2024

@author: mumble
"""

import PyPDF2
from openpyxl import Workbook

def extract_text_from_pdf(pdf_path):
    """
    Extract text from a PDF file.
    """
    text = ''
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        for page_num in range(len(pdf_reader.pages)):
            text += pdf_reader.pages[page_num].extract_text()
    return text

def write_to_excel(text, excel_path):
    """
    Write text data to an Excel file.
    """
    workbook = Workbook()
    sheet = workbook.active
    lines = text.split('\n')
    for row_num, line in enumerate(lines, start=1):
        sheet.cell(row=row_num, column=1, value=line)
    workbook.save(excel_path)

def main():
    pdf_path = 'my-resume-2021.pdf'  # Provide the path to your PDF file
    excel_path = 'output.xlsx'  # Provide the desired name for the output Excel file

    text = extract_text_from_pdf(pdf_path)
    write_to_excel(text, excel_path)
    print("Data extracted from PDF and written to Excel successfully!")

if __name__ == "__main__":
    main()

