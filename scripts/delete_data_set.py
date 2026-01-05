import os
from pathlib import Path

# Delete file
def delete_file(file: Path, directory: str='') -> None:
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)
        if directory == '':
            print(f"Deleted: {file}")
        else:
            print(f"Deleted: {file} in {directory}")

# Deletes all DOCX files in the STEGO data set directory
def delete_stego_data_set() -> None:
    stego_files = "data_set/stego_files"
    for directories in Path(stego_files).iterdir():
        for file in directories.iterdir():
            delete_file(file, directories.name)

# Deletes all DOCX files in the TEST_0 STEGO data set directory
def delete_TEST_0_data_set() -> None:
    test_0_data_set = "data_set/TEST_0"
    for file in Path(test_0_data_set).iterdir():
        delete_file(file, test_0_data_set)

# Deletes all DOCX files in the ATTACKED stego data set directory
def delete_attacked_data_set() -> None:
    attacked_stego_files = "data_set/attacked_stego_files"
    for attack_directories in Path(attacked_stego_files).iterdir():
        for stego_directories in attack_directories.iterdir():
            for file in stego_directories.iterdir():
                delete_file(file, stego_directories.name)

# Deletes all CSV results files in the results directory
def delete_results() -> None:
    results_folder = "results"
    for directories_analysis_type in Path(results_folder).iterdir():
        # Excel files directory is too difficult to manage, so Excel file deletion is skipped
        if directories_analysis_type.name != '0_excel_files':
            for directories_specific_results in directories_analysis_type.iterdir():
                for file in directories_specific_results.iterdir():
                    delete_file(file, directories_specific_results.name)

if __name__ == "__main__":
    # delete_stego_data_set()
    # delete_attacked_data_set()
    # delete_results()
    # delete_TEST_0_data_set()
    print("DELETED!")