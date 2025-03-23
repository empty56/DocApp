from utils import process_file
from FormatChecker.checkers import entry_checker, main_part_checker, tech_assignment_checker

def check_formatting_rules(file_path, document_part):
    if document_part == "entry":
        print("Checking entry formatting...")
        entry_checker.check_alignment(file_path)

    elif document_part == "tech_assignment":
        print("Checking technical assignment formatting...")
        tech_assignment_checker.check_alignment(file_path)

    elif document_part == "main_part":
        print("Checking main part formatting...")
        main_part_checker.check_alignment(file_path)

    # elif document_part == "additional_parts":
    #     print("Checking additional_parts formatting...")
    #     tech_assignment_checker.check_alignment(file_path)

    else:
        print("Unknown document part")


def main():
    # Get the DOCX file path
    # file_path = "D:/Diploma/Bakalavrat/Docs/TestDoc3.docx"
    # document_part = "tech_assignment"
    # process_file(file_path)
    # check_formatting_rules(file_path, document_part)
    file_path = "D:/Diploma/Bakalavrat/Docs/Main_part_copy4.docx"
    document_part = "main_part"
    process_file(file_path)
    check_formatting_rules(file_path, document_part)

if __name__ == "__main__":
    main()
