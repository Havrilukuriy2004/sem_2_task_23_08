import os
import re
from docx import Document

def replace_text_in_docx(file_path, regex, replacement):
    doc = Document(file_path)
    for paragraph in doc.paragraphs:
        if re.search(regex, paragraph.text, re.IGNORECASE):
            paragraph.text = re.sub(regex, replacement, paragraph.text, flags=re.IGNORECASE)
    doc.save(file_path)

def process_directory(directory, regex, replacement):
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.docx'):
                file_path = os.path.join(root, file)
                replace_text_in_docx(file_path, regex, replacement)

if __name__ == "__main__":
    directory = input("Enter the path to the directory: ")
    regex = input("Enter the regular expression: ")
    replacement = input("Enter the replacement text: ")
    process_directory(directory, regex, replacement)
