from fastapi import FastAPI, Query
from typing import List
import os
import docx2txt
import shutil

app = FastAPI()

@app.get('/get_and_insert_folder')
async def get_and_insert_folder(
    input_folder: str,
    output_folder: str,
    keywords: List[str] = Query(None)
):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    for filename in os.listdir(input_folder):
        if filename.endswith('.docx'):
            file_path = os.path.join(input_folder, filename)
            text = read_docx(file_path)
            if contains_keywords(text, keywords):
                dest_path = os.path.join(output_folder, filename)
                if not os.path.exists(dest_path):
                    try:
                        shutil.copy(file_path, dest_path)
                        print(f"Copied: {filename}")
                    except PermissionError as e:
                        print(f"PermissionError: {e}")
                    except Exception as e:
                        print(f"Error copying file {file_path} to {dest_path}: {e}")
                else:
                    print(f"File already exists and will not be copied again: {filename}")
    return {"status": "Completed"}

def read_docx(file_path):
    return docx2txt.process(file_path)

# Function to search for keywords in text
def contains_keywords(text, keywords):
    text_lower = text.lower()
    for keyword in keywords:
        if keyword.lower() in text_lower:
            return True
    return False
input_folder = r'C:\Users\Nagsrujan\Downloads\Model Discharge summaries'
output_folder = r'C:\Users\Nagsrujan\Downloads\TKA'
# keywords = list(input()) for dynamic input
keywords = ["Total knee Replacement", "TKR", "tkr"]
