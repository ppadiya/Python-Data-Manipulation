"""This Python script traverses a directory (including subfolders) and processes files (Word documents, Excel files, and PDFs) modified before a specified date. It converts older .doc files to .docx, adds a space to .docx and Excel files, and duplicates PDFs while updating their last modified timestamps. This is to prevent old files from auto-archiving or getting deleted """


import os
import datetime
from docx import Document
import openpyxl
from PyPDF2 import PdfReader, PdfWriter
import comtypes.client

# Specify the top-level directory where your documents are stored
top_directory_path = r'C:\Users\e057495\OneDrive - Mastercard'  # Replace with your actual path

# Get the target date (before which files should be processed)
target_date = datetime.datetime(2023, 1, 1, 0, 0, 0)  # Replace with your desired date

# Walk through the directory tree
for root, dirs, files in os.walk(top_directory_path):
    for filename in files:
        file_path = os.path.join(root, filename)

        # Check the last modified time of the file
        last_modified_time = datetime.datetime.utcfromtimestamp(os.path.getmtime(file_path))

        # Process only if the file was last modified before the target date
        if last_modified_time < target_date:
            try:
                # Check if the file is a Word document
                if filename.lower().endswith('.doc'):
                    try:
                        # Attempt to convert .doc to .docx using comtypes
                        word_app = comtypes.client.CreateObject("Word.Application")
                        doc = word_app.Documents.Open(file_path)
                        new_file_path = os.path.splitext(file_path)[0] + '.docx'
                        doc.SaveAs(new_file_path, FileFormat=16)  # 16 corresponds to .docx format
                        doc.Close()
                        word_app.Quit()

                        # Update the file_path to the new .docx file
                        file_path = new_file_path

                    except Exception as doc_conversion_error:
                        print(f"Error converting {filename}: {doc_conversion_error}")
                        continue

                # Check if the file is a Word document (.docx), Excel file (.xlsx or .xls), or PDF document
                if filename.lower().endswith(('.docx', '.xlsx', '.xls')):
                    # Use the appropriate library to process the file
                    if filename.lower().endswith('.docx'):
                        doc = Document(file_path)
                        doc.add_paragraph(' ')
                        doc.save(file_path)
                    elif filename.lower().endswith(('.xlsx', '.xls')):
                        try:
                            workbook = openpyxl.load_workbook(file_path)
                            workbook.save(file_path)
                        except Exception as excel_error:
                            print(f"Error processing {filename}: {excel_error}")

                # Check if the file is a PDF document
                elif filename.lower().endswith('.pdf'):
                    try:
                        with open(file_path, 'rb') as pdf_file:
                            pdf_reader = PdfReader(pdf_file)
                            pdf_writer = PdfWriter()

                            for page_num in range(len(pdf_reader.pages)):
                                pdf_writer.add_page(pdf_reader.pages[page_num])

                            with open(file_path, 'wb') as modified_pdf:
                                pdf_writer.write(modified_pdf)
                    except Exception as pdf_error:
                        print(f"Error processing {filename}: {pdf_error}")

            except Exception as general_error:
                print(f"Error processing {filename}: {general_error}")

print("Processing complete.")
