import os
import shutil

def organize_documents(root_folder):
    # Create separate folders for .docx, .doc, and .wps files
    docx_folder = os.path.join(root_folder, "DocxFiles")
    doc_folder = os.path.join(root_folder, "DocFiles")
    wps_folder = os.path.join(root_folder, "WpsFiles")

    os.makedirs(docx_folder, exist_ok=True)
    os.makedirs(doc_folder, exist_ok=True)
    os.makedirs(wps_folder, exist_ok=True)

    # Counters for each file type
    docx_count = 0
    doc_count = 0
    wps_count = 0

    # Recursively search through all folders and subfolders
    for foldername, subfolders, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.lower().endswith(".docx"):
                # Move .docx files to the DocxFiles folder
                src_path = os.path.join(foldername, filename)
                dst_path = os.path.join(docx_folder, filename)
                if not os.path.exists(dst_path):
                    shutil.move(src_path, dst_path)
                    print(f"Moved: {src_path} to {dst_path}")
                    docx_count += 1
                else:
                    print(f"File already exists at destination: {dst_path}. Skipping move.")
            elif filename.lower().endswith(".doc"):
                # Move .doc files to the DocFiles folder
                src_path = os.path.join(foldername, filename)
                dst_path = os.path.join(doc_folder, filename)
                if not os.path.exists(dst_path):
                    shutil.move(src_path, dst_path)
                    print(f"Moved: {src_path} to {dst_path}")
                    doc_count += 1
                else:
                    print(f"File already exists at destination: {dst_path}. Skipping move.")
            elif filename.lower().endswith(".wps"):
                # Move .wps files to the WpsFiles folder
                src_path = os.path.join(foldername, filename)
                dst_path = os.path.join(wps_folder, filename)
                if not os.path.exists(dst_path):
                    shutil.move(src_path, dst_path)
                    print(f"Moved: {src_path} to {dst_path}")
                    wps_count += 1
                else:
                    print(f"File already exists at destination: {dst_path}. Skipping move.")

    print("Organized .docx, .doc, and .wps files successfully!")
    print(f"Total .docx files moved: {docx_count}")
    print(f"Total .doc files moved: {doc_count}")
    print(f"Total .wps files moved: {wps_count}")

# Example usage: Replace 'your_root_folder' with the actual path
root_folder_path = r"C:\Users\Nagsrujan\Downloads\Dischargs for TKR"
organize_documents(root_folder_path)
