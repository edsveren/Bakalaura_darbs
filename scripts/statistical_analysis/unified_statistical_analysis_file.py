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
        result_folder: str,
        result_file: str, 
        data_to_csv: list[list],
        data_set_processed: bool,
        # append_starting_points: tuple,
        append_starting_point: int
    ) -> None:

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            if not data_set_processed:
                writer.writerow('') 
                for row in data_to_csv:
                    writer.writerow(row)
            else:
                # writer.writerow('')
                # for append_starting_point in append_starting_points: # THIS IS A HORRIBLE ATTEMPT
                writer.writerow(data_to_csv[append_starting_point]) # BETTER TO USE INT, FIX LATER!!!
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            for row in data_to_csv:
                writer.writerow(row)

def process_specific_data_set(
        data_set_processed: bool,
        data_to_csv: list[list],
        data_to_csv_list: list[list],
        append_starting_point: int
    ) -> list[list]:

    row_width = len(data_to_csv[append_starting_point])

    if not data_set_processed:
        cell_index = 0
        for cell in data_to_csv[1]:
            if cell == '':
                data_to_csv[1][cell_index] = '.'
            cell_index += 1
        data_to_csv_list.append(data_to_csv[1])
        if append_starting_point == 2:
            filler_row = ['.'] * row_width
            data_to_csv_list.append(filler_row)
        else:
            data_to_csv_list.append(data_to_csv[2])
    data_to_csv_list.append(data_to_csv[append_starting_point])
    return data_to_csv_list

# Export all individual statistical analysis data set results to a single CSV file
def export_specific_data_set_to_csv(
        result_file: str,
        data_to_csv_list: list[list]
        # existing_individual_stego_rows: list
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
            
        if existing_individual_stego_rows == []:
            for row in data_to_csv_list:
                writer.writerow(row[0:])
        else:
            for row_index in range(len(existing_individual_stego_rows)):
                writer.writerow(existing_individual_stego_rows[row_index] + data_to_csv_list[row_index][1:])
    
    print(f"Created a CSV file: {str(Path(result_file))}")

def clear_specific_data_set_to_csv() -> None:
    result_folder = f"results/statistical_analysis/all_specific_data_set_files"

    for file in Path(result_folder).iterdir():
        delete_file(file)

def print_output(data_to_csv: list[list]) -> None:
    # for item in data_to_csv:
    #     print(f"{item[0]}: {item[1]}")

    # Temporary solution
    for item in data_to_csv:
        label, *values = item
        print(f"{label}: {', '.join(map(str, values))}")
    print()

### Main function ###
def statistical_analysis(
        statistical_analysis_method: str, 
        analyzed_file: str|None, 
        statistical_analysis_method_execution: Callable[[Path, str, bool], tuple[list[list], int]]
    ) -> None:
    
    csv_path = f"results/statistical_analysis/{statistical_analysis_method}"
    if analyzed_file == None:
        csv_file = f"{csv_path}/all_files.csv"
    else:
        csv_file = f"{csv_path}/{analyzed_file}.csv"
    delete_file(Path(csv_file))

    clean_data_set = f'data_set/clean_files'
    stego_data_set = f'data_set/stego_files'

    data_sets = ["clean", "hide_in_text", "multilayer_hybrid", "two_bit_transformation", "modify_RGB_color_ch", "unispace", "unicode_homoglyphs"]

    for data_set_index in range(len(data_sets)):
        data_set_processed = False
        data_set = data_sets[data_set_index]
        if data_set == "clean":
            base_directory = clean_data_set
        else:
            base_directory = f"{stego_data_set}/stego_method_{data_set_index}"

        # Process only the chosen file
        if analyzed_file != None:
            docPath = Path(f"{base_directory}/{analyzed_file}.docx")
            if not docPath.is_file():
                print(f"File doesn't exist: {docPath}")
                print()
                data_set_index += 1
                continue

            print(f"Opened: {docPath}")
            data_to_csv, append_starting_point = statistical_analysis_method_execution(docPath, data_set, True)
            print_output(data_to_csv)
            export_to_csv(csv_path, csv_file, data_to_csv, False, append_starting_point)
        # Process all files
        else:
            # Check if the folder is empty
            has_docx_files = any(Path(base_directory).glob("*.docx"))
            if not has_docx_files:
                print(f"{str(Path(base_directory))} is empty!")
            else:
                # The result file containing all individual data set results
                result_file = f"results/statistical_analysis/all_specific_data_set_files/{data_set}.csv"
                data_to_csv_list = []

                for file in Path(base_directory).iterdir():
                    if file.is_file() and not file.name.startswith("."):    
                        print(f"Opened: {file}")
                        data_to_csv, append_starting_point = statistical_analysis_method_execution(file, data_set, False)

                        print_output(data_to_csv)
                        export_to_csv(csv_path, csv_file, data_to_csv, data_set_processed, append_starting_point)
                
                        if data_to_csv[0][1] == data_set:
                            data_to_csv_list = process_specific_data_set(data_set_processed, data_to_csv, data_to_csv_list, append_starting_point)
                        data_set_processed = True

                export_specific_data_set_to_csv(result_file, data_to_csv_list)

        print()

if __name__ == "__main__":
    clear_specific_data_set_to_csv()