import os

# Function to check suitable file type
def process_file(file_path):
    if not os.path.exists(file_path):
        return "File does not exist."

    extension = os.path.splitext(file_path)[1].lower()

    if extension == '.docx' or extension == '.doc':
        return "Processing your file..."
    else:
        return "Unsupported file format. Please use .docx or .doc files."
