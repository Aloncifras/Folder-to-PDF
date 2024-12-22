import os
import ctypes
from tkinter import filedialog, Tk
from docx2pdf import convert
from win32com import client

# Hide the root Tkinter window
root = Tk()
root.withdraw()

# Prompt user to select the root folder containing .doc and .docx files
ctypes.windll.user32.MessageBoxW(
    0,
    "Hit OK to select the root folder containing the MS Word DOC or DOCX files you wish to convert to a single PDF.",
    "Select Root Folder",
    0
)
root_folder = filedialog.askdirectory()

# Ensure a folder was selected
if not root_folder:
    ctypes.windll.user32.MessageBoxW(0, "You must select a root folder.", "Error", 0)
    exit()

# Get the folder name for naming the final PDF
folder_name = os.path.basename(root_folder.rstrip("/\\"))

# Determine output PDF path
output_pdf = os.path.join(root_folder, f"{folder_name}.pdf")
counter = 1
while os.path.exists(output_pdf):
    output_pdf = os.path.join(root_folder, f"{folder_name} ({counter}).pdf")
    counter += 1

# Temporary file for merged Word document
merged_doc_path = os.path.join(root_folder, "MergedDocument.docx")

# Check if the merged document already exists
if os.path.exists(merged_doc_path):
    os.remove(merged_doc_path)

# Merge Word documents into a single Word file while preserving formatting
def merge_docs_with_com(input_folder, output_path):
    word = client.Dispatch("Word.Application")
    word.Visible = False
    merged_doc = word.Documents.Add()

    for root, _, files in os.walk(input_folder):
        for file in sorted(files):  # Sort files to maintain a predictable order
            if file.lower().endswith((".doc", ".docx")):
                doc_path = os.path.join(root, file)
                doc_path = os.path.abspath(doc_path)  # Ensure absolute path
                try:
                    if not os.path.exists(doc_path):
                        print(f"File not found: {doc_path}")
                        continue
                    print(f"Attempting to open: {doc_path}")  # Debugging log
                    sub_doc = word.Documents.Open(doc_path)

                    # Insert content of the document directly at the end of the merged document
                    merged_doc.Range(merged_doc.Content.End - 1).InsertFile(doc_path)

                    # Add a page break after each document
                    merged_doc.Range(merged_doc.Content.End - 1).InsertBreak(7)  # 7 corresponds to wdPageBreak

                    sub_doc.Close(False)
                except Exception as e:
                    print(f"Failed to merge {doc_path}: {e}")

    merged_doc.SaveAs2(output_path)
    merged_doc.Close()
    word.Quit()

# Merge all Word documents into a single Word file
merge_docs_with_com(root_folder, merged_doc_path)

# Convert the merged Word document to PDF
try:
    convert(merged_doc_path)
    os.rename(merged_doc_path.replace(".docx", ".pdf"), output_pdf)
except Exception as e:
    print(f"Failed to convert merged document to PDF: {e}")

# Cleanup temporary merged document
# if os.path.exists(merged_doc_path):
#     os.remove(merged_doc_path)

# Notify user of completion
ctypes.windll.user32.MessageBoxW(0, f"Conversion complete. The final PDF is saved at:\n{output_pdf}", "Done", 0)
