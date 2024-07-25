import os
import shutil
import docx2txt


# Function to read text from a .docx file
def read_docx(file_path):
    return docx2txt.process(file_path)


# Function to search for keywords in text
def contains_keywords(text, keywords):
    text_lower = text.lower()
    for keyword in keywords:
        if keyword.lower() in text_lower:
            return True
    return False


# Main function to process the folder
def get_and_insert_folder(input_folder, output_folder, keywords):
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


# Define the input folder, output folder, and keywords
input_folder = r'C:\Users\Nagsrujan\Downloads\Model Discharge summaries'
output_folder = r'C:\Users\Nagsrujan\Downloads\TKA'
# keywords = list(input()) for dynamic input
keywords = ["Total knee Replacement", "TKR", "tkr"]

# Run the script
get_and_insert_folder(input_folder, output_folder, keywords)
