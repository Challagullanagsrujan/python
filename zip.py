import os
import shutil
from zipfile import ZipFile


def separate_and_zip_files(source_directory):
    docx_folder = os.path.join(source_directory, 'docx_files')
    doc_folder = os.path.join(source_directory, 'doc_files')

    # Create directories if they don't exist
    os.makedirs(docx_folder, exist_ok=True)
    os.makedirs(doc_folder, exist_ok=True)

    # Walk through all files and subdirectories in the source directory
    for root, dirs, files in os.walk(source_directory):
        for file in files:
            file_path = os.path.join(root, file)

            # Skip files in the target folders
            if root.startswith(docx_folder) or root.startswith(doc_folder):
                continue

            if file.endswith('.docx'):
                destination = os.path.join(docx_folder, file)
                if os.path.exists(destination):
                    print(f"File already exists: {destination}. Skipping move.")
                else:
                    shutil.move(file_path, destination)
                    print(f"Moved: {file_path} to {destination}")
            elif file.endswith('.doc'):
                destination = os.path.join(doc_folder, file)
                if os.path.exists(destination):
                    print(f"File already exists: {destination}. Skipping move.")
                else:
                    shutil.move(file_path, destination)
                    print(f"Moved: {file_path} to {destination}")

    # Zip the folders
    zip_folder(docx_folder, os.path.join(source_directory, 'docx_files.zip'))
    zip_folder(doc_folder, os.path.join(source_directory, 'doc_files.zip'))


def zip_folder(folder_path, zip_path):
    with ZipFile(zip_path, 'w') as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)
                print(f"Added to zip: {file_path} as {arcname}")


# Example usage
source_directory = r"C:\Users\Nagsrujan\Downloads\Dischargs for TKR"  # Replace with your folder path
separate_and_zip_files(source_directory)
