import os
import win32com.client

docx_path = os.path.abspath("documentation/License_Plate_Recognition_Project_Report.docx")
pdf_path = os.path.abspath("documentation/License_Plate_Recognition_Project_Report.pdf")

print(f"Converting '{docx_path}' to '{pdf_path}' via Word COM...")

word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(docx_path)
    # FileFormat=17 is wdFormatPDF
    doc.SaveAs(pdf_path, FileFormat=17)
    doc.Close()
    print("PDF export successful!")
except Exception as e:
    print(f"Error converting to PDF: {e}")
finally:
    word.Quit()

if os.path.exists(pdf_path):
    print(f"PDF verified on disk! Size: {os.path.getsize(pdf_path)} bytes")
