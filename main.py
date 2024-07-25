import os
import subprocess


def merge_docx_files_with_pandoc(docx_files, output_file):
    # Join the list of docx files into a single string separated by spaces
    input_files_str = ' '.join([f'"{file}"' for file in docx_files])

    # Create the pandoc command
    pandoc_command = f'pandoc {input_files_str} -o "{output_file}"'

    # Run the pandoc command
    result = subprocess.run(pandoc_command, shell=True, capture_output=True, text=True)

    if result.returncode != 0:
        print("An error occurred during merging with pandoc:")
        print(result.stderr)
    else:
        print(f"Merged document created at {output_file}")


# Directory containing the .docx files
directory = os.path.abspath("docxFiles")

# Verify the directory exists
if not os.path.exists(directory):
    raise FileNotFoundError(f"The directory '{directory}' does not exist.")

# List of .docx files to merge
docx_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.docx')]

# Check if there are files to merge
if not docx_files:
    raise ValueError("No .docx files found in the directory.")

# Output file
output_file = os.path.join(directory, "SGICM_LogisticaQR.docx")

# Merge the .docx files
merge_docx_files_with_pandoc(docx_files, output_file)
