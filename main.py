import os
import re
from docx import Document

def convert_doc_to_txt(input_dir, output_dir):
    # Check if the output directory exists, create it if not
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    print(f"Converting .doc files in {input_dir} to .txt files in {output_dir}")
    # Get a list of all files in the input directory
    files = os.listdir(input_dir)

    # Iterate over the files and convert .doc files to .txt
    for file in files:
        if file.endswith('.doc') or file.endswith('.docx'):
            # Open the .doc file
            doc_path = os.path.join(input_dir, file)
            doc = Document(doc_path)

            # Create the output .txt file path
            txt_filename = os.path.splitext(file)[0] + '.txt'
            txt_path = os.path.join(output_dir, txt_filename)

            # Extract the text from the .doc file and save it as .txt
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                for paragraph in doc.paragraphs:
                    txt_file.write(paragraph.text)
                    txt_file.write('\n')
    
            print(f"Converted {file} to {txt_filename}")
        
        elif file.endswith('.txt'):
            # Open the .txt file
            txt_path = os.path.join(input_dir, file)
            txt_file = open(txt_path, 'r', encoding='utf-8')
            txt = txt_file.read()

            # Create the output .txt file path
            txt_filename = os.path.splitext(file)[0] + '.txt'
            txt_path = os.path.join(output_dir, txt_filename)

            # Extract the text from the .txt file and save it as .txt
            with open(txt_path, 'w', encoding='utf-8') as txt_file:
                txt_file.write(txt)
    
            print(f"Converted {file} to {txt_filename}")

# Example usage
input_directory = "C:\\Users\\MSI GL65 Leopard\\Downloads\\input\\Test Cases"
output_directory = "C:\\Users\\MSI GL65 Leopard\\Downloads\\test"
convert_doc_to_txt(input_directory, output_directory)
