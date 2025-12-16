import csv
import time
from pathlib import Path
from typing import Callable
from itertools import zip_longest
import win32com.client as win32
from scripts.delete_data_set import delete_file

# Export time of each active steganalysis attack to CSV
def export_attack_time_to_csv_individual(
        attack_type: str,
        data_to_csv: list[list]
          ) -> None:
    
    # Delete previous CSV results
    clean_results_individual(attack_type)
    
    result_file = f"results/active_attacks/not_transposed/{attack_type}.csv"
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
        print(f"Created a CSV file: {str(Path(result_file))}")
    
    # Transpose the resulting CSV file
    csv_transpose(attack_type)

# Transpose the CSV file
def csv_transpose(attack_type: str):
    input_file = f"results/active_attacks/not_transposed/{attack_type}.csv"
    output_file_ = f"results/active_attacks/transposed/{attack_type}_transposed.csv"
    
    # Read CSV rows
    with open(input_file, newline="", encoding="utf-8") as input_file:
        rows = list(csv.reader(input_file, delimiter=";"))

    # Transpose using zip_longest to handle unequal row and columns lengths
    # Replace "missing cells" with empty strings
    transposed_rows = list(zip_longest(*rows, fillvalue=""))

    # Write transposed rows
    with open(output_file_, "w", newline="", encoding="utf-8") as output_file:
        writer = csv.writer(output_file, delimiter=";")
        for row in transposed_rows:
            new_row = ["", *row]
            writer.writerow(new_row)
    
    print(f"Created a CSV file: {str(Path(output_file_))}")

# Export time of all active steganalysis attacks into a single CSV file
def export_attack_time_to_csv_all(transposed_state: str, transposed_or_not_type: str) -> None:
    base = f"results/active_attacks/{transposed_state}"
    all_attacks = f"all_attacks{transposed_or_not_type}.csv"
    result_file = f"{base}/{all_attacks}"
    with open(result_file, "w", encoding="utf-8", newline="") as output_file:
        writer = csv.writer(output_file, delimiter=";")
        for attack_result_files in Path(base).iterdir():
            if attack_result_files.is_file() and not attack_result_files.name.startswith(".") and attack_result_files.name != all_attacks:
                with open(attack_result_files, "r", encoding="utf-8") as input_file:
                    reader = csv.reader(input_file, delimiter=";")
                    for row in reader:
                        writer.writerow(row)
                writer.writerow('')  # Add an empty line between different attack results
    
    print(f"Created a CSV file: {str(Path(result_file))}")

# Clean individual active steganalysis attack results CSV file
def clean_results_individual(attack_type: str) -> None:
    file = Path(f"results/active_attacks/not_transposed/{attack_type}.csv")
    file_transposed = Path(f"results/active_attacks/transposed/{attack_type}_transposed.csv")
    delete_file(file)
    delete_file(file_transposed)

# Clean the unified active steganalysis attack results CSV file
def clean_results_all_attacks() -> None:
    file = Path("results/active_attacks/not_transposed/all_attacks.csv")
    file_transposed = Path("results/active_attacks/transposed/all_attacks_transposed.csv")
    delete_file(file)
    delete_file(file_transposed)

# Clean every single active steganalysis attack results file
def clean_results_folder() -> None:
    base = "results/active_attacks"
    for folders in Path(base).iterdir():
        for attack_result_files in folders.iterdir():
            delete_file(attack_result_files)

# Process and display active steganalysis attack time
def time_display(
        timelapse_object_name: str, 
        timeLapse: float
        ) -> float:
    timeLapseInt = int(timeLapse)
    timeLapseSeconds = int(timeLapse * 100) / 100
    timeLapseMinutes = timeLapseInt // 60
    timeLapseSecInMin = timeLapseInt % 60

    timeLapseFloat = float(f"{timeLapseMinutes}.{timeLapseSecInMin:02d}")

    print(f"{timelapse_object_name} timelapse (s): {timeLapseSeconds}")
    print(f"{timelapse_object_name} timelapse (min): {timeLapseFloat}")
    return timeLapseFloat

### A unified active steganalysis attack framework ###
def unified_attack(
        attack_type: str, 
        attack: Callable[..., None],
        UsingWin: bool
        ) -> None:
    # Stego-files' location
    stego_base = "data_set/stego_files"
    # Attacked stego-files' saving location
    attacked_base = "data_set/attacked_stego_files"

    # Setup time logging variables
    totalTimeLapse = 0
    directoryTimeLapseList = []
    data_set_type_list = []

    # MS Word PDF format index
    wdFormatPDF = 17 
    # MS Word default (DOCX) document file format index
    wdFormatDocumentDefault = 16
    # Remove all document information index
    wdRDIAll = 99
    # Opening MS Word is disabled by default
    word = None

    # If attacking using MS Word, create an instance
    if UsingWin:
        word = win32.Dispatch("Word.Application")

        # Process the DOCX file without visibly opening MS Word
        word.Visible = False
        word.DisplayAlerts = 0

        # Close any open documents without saving
        while word.Documents.Count > 0:
            word.Documents(1).Close(SaveChanges=0)

    # Looping through each stego-method file folder
    for directories in Path(stego_base).iterdir():
        directoryTimeLapse = 0

        # Looping through each stego-file
        for file in directories.iterdir():
            print()
            # Ignore temporary and git files
            if file.name.startswith('~$') or file.name.startswith('.'):
                continue
            start = time.time()

            # Stego-file
            stegoDocPath = str(Path(f"{stego_base}/{directories.name}/{file.name}").resolve())
            stegoPDFPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.stem}.pdf").resolve())
            attackedStegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}").resolve())
            
            print(f"Attacking: {stegoDocPath}")

            # If attacking using MS Word
            if UsingWin:
                match attack_type:
                    case "05_impersonation_attack":
                        attack(word, stegoDocPath, stegoPDFPath, attackedStegoDocPath, wdFormatPDF, wdFormatDocumentDefault)
                    case "10_document_inspector_attack":
                        attack(word, stegoDocPath, attackedStegoDocPath, wdFormatDocumentDefault, wdRDIAll)
                    case _:
                        attack(word, stegoDocPath, attackedStegoDocPath, wdFormatDocumentDefault)
            else:
                attack(stegoDocPath, attackedStegoDocPath)
            print(f"Saved: {attackedStegoDocPath}")

            end = time.time()
            fileTimeLapse = end - start
            directoryTimeLapse += fileTimeLapse
            print(f"Time taken (s): {int(fileTimeLapse * 100) / 100}")
            print()

        directoryTimeLapseFloat = time_display(f"{directories.name} directory", directoryTimeLapse)

        data_set_type_list.append(directories.name)
        directoryTimeLapseList.append(directoryTimeLapseFloat)
        totalTimeLapse += directoryTimeLapse
    print()
    
    # If attacking using MS Word, exit
    if UsingWin:
        word.Quit()

    totalTimeLapseFloat = time_display("Total", totalTimeLapse)

    # Data to save to CSV
    data_to_csv = [
        ["Active stego-attack type", attack_type],
        # ["Total attack time lapse (s)", totalTimeLapseSec],
        ["Total attack time lapse (min)", totalTimeLapseFloat],
        ["Stego-file data set", *data_set_type_list], 
        ["Attack time for each stego-file data set (min)", *directoryTimeLapseList]
    ]

    export_attack_time_to_csv_individual(attack_type, data_to_csv)

if __name__ == "__main__":
    # clean_results_folder()
    clean_results_all_attacks()
    export_attack_time_to_csv_all("not_transposed", "")
    # csv_transpose("all_attacks")
    export_attack_time_to_csv_all("transposed", "_transposed")