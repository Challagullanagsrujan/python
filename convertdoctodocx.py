import os
import re
import win32com.client as win32
from win32com.client import constants

def save_as_docx(word, path):
    # Open the .doc file
    doc = word.Documents.Open(path)
    doc.Activate()

    # Rename path with .docx
    new_file_abs = os.path.abspath(path)
    new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)

    # Save and Close
    word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
    doc.Close(False)

    print(f'Converted {path} to {new_file_abs}')

def convert_folder_docs_to_docx(folder_path):
    # Opening MS Word
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False

    # Create a list of paths to .doc files
    paths = [os.path.join(folder_path, filename) for filename in os.listdir(folder_path) if filename.lower().endswith('.doc')]

    for path in paths:
        try:
            count=count+1
            save_as_docx(word, path)
        except Exception as e:
            print(f'Failed to convert {path}: {e}')
    print(count)

    # Quit Word application
    word.Quit()

# Specify the folder containing .doc files
folder_path = r"C:\Users\Nagsrujan\Downloads\Dischargs for TKR (2)\Dischargs for TKR\DocFiles"

convert_folder_docs_to_docx(folder_path)
