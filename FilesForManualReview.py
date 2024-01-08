import os
import shutil
import fitz  # Import PyMuPDF for handling PDF files
from docx import Document  # Import python-docx for handling Word documents

def create_search_folder(dst_folder, search_word):
    """
    Creates a folder for each search word within the destination directory.
    """
    search_folder = os.path.join(dst_folder, search_word)
    os.makedirs(search_folder, exist_ok=True)  # Create the folder if it does not exist
    return search_folder

def search_text_in_pdf(pdf_path, search_words):
    """
    Searches for the presence of any search words within a PDF file.
    """
    with fitz.open(pdf_path) as pdf:
        for page in pdf:
            text = page.get_text().lower()  # Extract text and convert to lowercase
            if any(word.lower() in text for word in search_words):
                return True  # Return True if any search word is found
    return False  # Return False if no search words are found

def search_text_in_docx(docx_path, search_words):
    """
    Searches for the presence of any search words within a Word document.
    """
    doc = Document(docx_path)
    for para in doc.paragraphs:
        if any(word.lower() in para.text.lower() for word in search_words):
            return True  # Return True if any search word is found
    return False  # Return False if no search words are found

def find_and_copy_files(src_folder, dst_folder, file_ext, search_words):
    """
    Finds files that match the search criteria and copies them to the destination folder.
    """
    files_to_copy = []  # List to hold files that meet the search criteria
    for root, dirs, files in os.walk(src_folder):
        for file in files:
            src_path = os.path.join(root, file)
            for search_word in search_words:
                if file.endswith('.pdf') and search_text_in_pdf(src_path, search_words):
                    dst_path = os.path.join(create_search_folder(dst_folder, search_word), file)
                    files_to_copy.append((src_path, dst_path))
                elif file.endswith('.docx') and search_text_in_docx(src_path, search_words):
                    dst_path = os.path.join(create_search_folder(dst_folder, search_word), file)
                    files_to_copy.append((src_path, dst_path))
                elif file.endswith('.vsd'):
                    # For Visio files, simply check if the search word is in the file name
                    dst_path = os.path.join(create_search_folder(dst_folder, search_word), file)
                    files_to_copy.append((src_path, dst_path))
                    break  # Stop searching if we've found a match to avoid duplicate copies
    return files_to_copy

def copy_files(file_paths):
    """
    Copies files from the source path to the destination path.
    """
    for src_path, dst_path in file_paths:
        shutil.copy(src_path, dst_path)  # Perform the file copy

# Main execution
if __name__ == "__main__":
    # Define the source and destination folders and search words
    src_folder = "path/to/shared/drive"  # Replace with the actual source folder path
    dst_folder = "path/to/destination"  # Replace with the actual destination folder path
    file_ext = [".docx", ".pdf", ".vsd"]  # List of file extensions to search for
    search_words = ["searchword1", "searchword2", "searchword3"]  # List of search words

    # Find files that match the search criteria and copy them
    files_to_copy = find_and_copy_files(src_folder, dst_folder, file_ext, search_words)
    copy_files(files_to_copy)