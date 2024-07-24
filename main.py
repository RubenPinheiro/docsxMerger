import os
from docx import Document
from docxcompose.composer import Composer

def merge_docx(files, output_file):
    if not files:
        raise ValueError("No files to merge.")

    # Open the first document
    master = Document(files[0])
    composer = Composer(master)

    # Append the rest of the documents
    for file in files[1:]:
        doc = Document(file)
        composer.append(doc)

    # Save the merged document
    composer.save(output_file)

# Directory containing the .docx files
directory = os.path.abspath("docxFiles")

# Verify the directory exists
if not os.path.exists(directory):
    raise FileNotFoundError(f"The directory '{directory}' does not exist.")

# List of .docx files to merge
docx_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.docx')]

# Check if there are .docx files to merge
if not docx_files:
    raise FileNotFoundError("No .docx files found in the directory.")

# Output file
output_file = os.path.join(directory, "SGICM_LogisticaQR.docx")

# Merge the .docx files
merge_docx(docx_files, output_file)

print(f"All files merged into {output_file}")
