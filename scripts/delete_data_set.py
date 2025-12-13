import os
from pathlib import Path

# Delete file
def delete_file(file: Path, directory: str) -> None:
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)
        print(f"Deleted: {file} in {directory}")

# Deletes all DOCX files in the STEGO data set directory
def delete_stego_data_set():
    stego_files = "data_set/stego_files"
    for directories in Path(stego_files).iterdir():
        for file in directories.iterdir():
            #docPath = f"{stego_files}/{directories.name}/{file.name}"
            delete_file(file, directories.name)

# Deletes all DOCX files in the ATTACKED stego data set directory
def delete_attacked_data_set():
    attacked_stego_files = "data_set/attacked_stego_files"
    for attack_directories in Path(attacked_stego_files).iterdir():
        for stego_directories in attack_directories.iterdir():
            for file in stego_directories.iterdir():
                #docPath = f"{attacked_stego_files}/{stego_directories.name}/{file.name}"
                delete_file(file, stego_directories.name)

# Deletes all CSV results files in the results directory
def delete_results():
    results_folder = "results"
    for directories_analysis_type in Path(results_folder).iterdir():
        # Excel files directory is too difficult to manage, so we skip deleting those files
        if directories_analysis_type.name != 'excel_files':
            for directories_specific_results in directories_analysis_type.iterdir():
                for file in directories_specific_results.iterdir():
                    delete_file(file, directories_specific_results.name)

if __name__ == "__main__":
    delete_stego_data_set()
    delete_attacked_data_set()
    delete_results()
    print("DELETED!")