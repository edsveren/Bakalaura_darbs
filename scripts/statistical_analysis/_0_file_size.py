import os
from pathlib import Path
import scripts.statistical_analysis.unified_statistical_analysis_file as unified_statistical_analysis_file

# Count number of documents in the data set
def count_number_of_documents_in_data_set(path: Path) -> int:
    data_set_directory = path.parent
    count = 0
    for file in data_set_directory.iterdir():
        if file.is_file() and file.suffix == '.docx':
            count += 1
    return count

# Get current file sizes
def get_file_size(path: Path) -> tuple[int, float, float]:
    file_size_bytes = os.path.getsize(path)
    file_size_kilobytes = round(file_size_bytes/1024, 2)
    file_size_megabytes = round(file_size_kilobytes/1024, 2)
    return file_size_bytes, file_size_kilobytes, file_size_megabytes

### Main function ###
def file_size_analysis(path: Path, data_set: str, chosen_file: bool) -> tuple[list[list], int]:
    
    data_set_document_amount = count_number_of_documents_in_data_set(path)
    file_size_bytes, file_size_kilobytes, file_size_megabytes = get_file_size(path)

    # Data export
    if not chosen_file:
        data_to_csv = [
            ["Data set", data_set],
            ["Document Name", "Data set document amount", "Document file size (B)", "Document file size (KB)", "Document file size (MB)"],
            [path.stem, data_set_document_amount, file_size_bytes, file_size_kilobytes, file_size_megabytes]
        ]
    else:
        data_to_csv = [
            ["Document Name", path.stem],
            ["Data set", data_set],
            ["Data set document amount", data_set_document_amount],
            ["Document file size (B)", file_size_bytes],
            ["Document file size (KB)", file_size_kilobytes],
            ["Document file size (MB)", file_size_megabytes]
        ]
    return data_to_csv, 2

def main() -> None:
    unified_statistical_analysis_file.statistical_analysis('0_file_size', 'TEST_0', file_size_analysis)
    unified_statistical_analysis_file.statistical_analysis('0_file_size', None, file_size_analysis)

if __name__ == "__main__":
    main()