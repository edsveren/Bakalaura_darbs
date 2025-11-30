import os
import csv
from pathlib import Path
from typing import Callable

# Delete file
def delete_file(file: Path) -> None:
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)
        print(f"Deleted: {str(file)}")

def export_to_csv(
        result_file: str, 
        data_to_csv: list[list]
    ) -> None:

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            for row in data_to_csv:
                writer.writerow(row)
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            for row in data_to_csv:
                writer.writerow(row)
            writer.writerow('')

def print_output(data_to_csv: list[list]) -> None:
    # for item in data_to_csv:
    #     print(f"{item[0]}: {item[1]}")

    # Temporary solution
    for item in data_to_csv:
        label, *values = item
        print(f"{label}: {', '.join(map(str, values))}")
    print()

def singular_check(
        statistical_analysis_method: str, 
        analyzed_file: str|None, 
        statistical_analysis_method_execution: Callable[[Path, str, bool], list[list]]
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
            data_to_csv = statistical_analysis_method_execution(docPath, data_set, True)
            print_output(data_to_csv)
            export_to_csv(csv_file, data_to_csv)
        # Process all files
        else:
            # Check if the folder is empty
            has_docx_files = any(Path(base_directory).glob("*.docx"))
            if not has_docx_files:
                print(f"{str(Path(base_directory))} is empty!")
            else:
                for file in Path(base_directory).iterdir():
                    if file.is_file() and not file.name.startswith("."):    
                        print(f"Opened: {file}")
                        data_to_csv = statistical_analysis_method_execution(file, data_set, False)         
                        # print_output(data_to_csv)
        print()
        
        