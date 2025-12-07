import os
import csv
from pathlib import Path
from typing import Callable

# Delete file
def delete_file(file: Path) -> None:
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)
        print(f"Deleted: {str(file)}")

# Export results to a CSV file
def export_to_csv(
        result_file: str, 
        data_to_csv: list[list],
        data_set_processed: bool,
        append_starting_point: int
    ) -> None:

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            # If statistical analysis results weren't yet exported
            # Add the header
            if not data_set_processed:
                writer.writerow('') 
                for row in data_to_csv:
                    writer.writerow(row)
            # Otherwise, export just the raw data
            else:
                writer.writerow(data_to_csv[append_starting_point])
    else:
        # File wasn't yet created, export everything
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            for row in data_to_csv:
                writer.writerow(row)

# Process all the individual statistical analysis results
# For a specific data set to export to a unified CSV file  
def process_specific_data_set(
        data_set_processed: bool,
        data_to_csv: list[list],
        data_to_csv_list: list[list],
        append_starting_point: int
    ) -> list[list]:

    # Get the longest row width
    row_width = len(data_to_csv[append_starting_point])

    # If results were not yet processed, adjust headers
    if not data_set_processed:
        # For better export, add '.' to all empty cells
        cell_index = 0
        for cell in data_to_csv[1]:
            if cell == '':
                data_to_csv[1][cell_index] = '.'
            cell_index += 1
        data_to_csv_list.append(data_to_csv[1])
        
        # As it stands currently, all headers are up to 2 or 3 rows long
        # So when it's the former, fill the 3rd row with filler for better export
        if append_starting_point == 2:
            filler_row = ['.'] * row_width
            data_to_csv_list.append(filler_row)
        # Otherwise just export the latter case
        else:
            data_to_csv_list.append(data_to_csv[2])
    # Last, get the raw data row itself
    data_to_csv_list.append(data_to_csv[append_starting_point])
    return data_to_csv_list

# Export all single data set statistical analysis results to a unified CSV file
def export_specific_data_set_to_csv(
        result_file: str,
        data_to_csv_list: list[list]
    ) -> None:

    # Read existing statistical analysis data set result file data if it exists
    if Path(result_file).is_file():
        with open(result_file, "r", encoding="utf-8", newline="") as input_file:
            reader = csv.reader(input_file, delimiter=";")
            existing_individual_stego_rows = list(reader)
    # Otherwise, create an empty list
    else:
        existing_individual_stego_rows = []

    # Append individual statistical analysis data set results to the unified data set result file
    with open(result_file, "w", encoding="utf-8", newline="") as output_file:
        writer = csv.writer(output_file, delimiter=";")
            
        # If the result file doesn't exist, export everything    
        if existing_individual_stego_rows == []:
            for row in data_to_csv_list:
                writer.writerow(row)
        # Otherwise, append the previous columns and do not export the document name (redundant)
        else:
            for row_index in range(len(existing_individual_stego_rows)):
                writer.writerow(existing_individual_stego_rows[row_index] + data_to_csv_list[row_index][1:])
    
    print(f"Created a CSV file: {str(Path(result_file))}")

# Clear all single data set statistical analysis results
def clear_specific_data_set_to_csv() -> None:
    result_folder = f"results/statistical_analysis/all_specific_data_set_files"

    for file in Path(result_folder).iterdir():
        delete_file(file)

# Export all statistical analysis individual data set results to a single CSV file
def export_to_csv_all() -> None:

    # The result directory statistical analysis individual data set results
    csv_path = "results/statistical_analysis/all_specific_data_set_files"
    # Output CSV file
    result_file = f"{csv_path}/all_data_set_unified_file.csv"

    # Delete existing result file if exists
    delete_file(Path(result_file))
    result_file_exists = False

    # Loop through each statistical analysis individual data set result file and merge them
    for file in Path(csv_path).iterdir():

        # Access the statistical analysis individual data set result file
        if file.is_file() and file != result_file and not file.name.startswith("."):
            # Read the file data and
            # Append the statistical analysis results to the result file
            with open(file, "r", encoding="utf-8", newline="") as input_file, open(result_file, "a+", encoding="utf-8", newline="") as output_file:
                reader = csv.reader(input_file, delimiter=";")
                writer = csv.writer(output_file, delimiter=";")

                # Add a separation row after the first table
                if result_file_exists:
                    writer.writerow('')
                row_header_index = 0
                for row in reader:
                    # No longer write first two rows
                    if row_header_index == 2:
                        result_file_exists = True
                    
                    # Write the first header row
                    if not result_file_exists and row_header_index < 1:
                        writer.writerow(['Data set'] + row)
                    elif row_header_index == 2:
                    # Write data set name
                        writer.writerow([file.stem] + row)
                    # Ignore headers if result file already exists
                    elif not result_file_exists or row_header_index >= 2:
                        writer.writerow([''] + row)
                    row_header_index += 1

    print(f"Created a CSV file: {str(Path(result_file))}")

# Print results to terminal
def print_output(data_to_csv: list[list]) -> None:
    for item in data_to_csv:
        header, *values = item
        print(f"{header}: {', '.join(map(str, values))}")
    print()

### Main function ###
def statistical_analysis(
        statistical_analysis_method: str, 
        analyzed_file: str|None, 
        statistical_analysis_method_execution: Callable[[Path, str, bool], tuple[list[list], int]]
    ) -> None:
    
    # Create a CSV result path depending on if a single file is specified or not
    csv_path = f"results/statistical_analysis/{statistical_analysis_method}"
    if analyzed_file == None:
        csv_file = f"{csv_path}/all_files.csv"
    else:
        csv_file = f"{csv_path}/{analyzed_file}.csv"
    delete_file(Path(csv_file))

    clean_data_set = f'data_set/clean_files'
    stego_data_set = f'data_set/stego_files'

    data_sets = ["clean", "hide_in_text", "multilayer_hybrid", "two_bit_transformation", "modify_RGB_color_ch", "unispace", "unicode_homoglyphs"]

    # Loop through each data set
    for data_set_index in range(len(data_sets)):
        # Create a boolean for the purposes of tracking
        # Whether the header should be added or not
        data_set_processed = False
        data_set = data_sets[data_set_index]

        # Choose either clean or stego-data set directory
        if data_set == "clean":
            base_directory = clean_data_set
        else:
            base_directory = f"{stego_data_set}/stego_method_{data_set_index}"

        # Process only the chosen file
        if analyzed_file != None:
            # Document path in the data set
            docPath = Path(f"{base_directory}/{analyzed_file}.docx")

            # If the data set doesn't exist, move on
            if not docPath.is_file():
                print(f"File doesn't exist: {docPath}")
                print()
                data_set_index += 1
                continue

            # Get data results and the raw data index from the specific statistical analysis method
            print(f"Opened: {docPath}")
            data_to_csv, append_starting_point = statistical_analysis_method_execution(docPath, data_set, True)
            print_output(data_to_csv)
            export_to_csv(csv_file, data_to_csv, False, append_starting_point)
        # Process all files
        else:
            # Check if the folder is empty
            has_docx_files = any(Path(base_directory).glob("*.docx"))
            if not has_docx_files:
                print(f"{str(Path(base_directory))} is empty!")
            else:
                # The result file containing all individual data set results
                result_file = f"results/statistical_analysis/all_specific_data_set_files/{data_set}.csv"
                # Create an empty list to add each data set row to export a specific data set statistical analysis results
                data_to_csv_list = []

                # Loop through each file in the data set directory
                for file in Path(base_directory).iterdir():
                    # Find the non-git documents
                    if file.is_file() and not file.name.startswith("."):
                        # Get data results and the raw data index from the specific statistical analysis method
                        print(f"Opened: {file}")
                        data_to_csv, append_starting_point = statistical_analysis_method_execution(file, data_set, False)
                        print_output(data_to_csv)
                        export_to_csv(csv_file, data_to_csv, data_set_processed, append_starting_point)
                        
                        # If document belongs to the currently iterated data set,
                        # Process its statistical analysis results for export to a unified data set CSV file
                        if data_to_csv[0][1] == data_set:
                            data_to_csv_list = process_specific_data_set(data_set_processed, data_to_csv, data_to_csv_list, append_starting_point)
                        data_set_processed = True
                # Export all current data set results to a single unified CSV file
                export_specific_data_set_to_csv(result_file, data_to_csv_list)
        print()

if __name__ == "__main__":
    export_to_csv_all()
    # clear_specific_data_set_to_csv()