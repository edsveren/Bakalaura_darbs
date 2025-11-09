import os
import csv
import time
from pathlib import Path
from typing import Callable
from docx.document import Document as DocumentObject

def log_time_attack_to_csv_individual(
          attack_type: str, 
          totalTimeLapseSec: float, 
          totalTimeLapseMin: float, 
          data_set_type_list: list, 
          directoryTimeLapseList: list
          ) -> None:
    result_file = f"results/active_attacks_logs/{attack_type}.csv"
    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
                writer = csv.writer(output_file, delimiter=";")
                writer.writerow(["Active stego-attack type", attack_type])
                writer.writerow(["Total attack time lapse (s)", totalTimeLapseSec])
                writer.writerow(["Total attack time lapse (min)", totalTimeLapseMin])
                writer.writerow(["Stego-file data set", *data_set_type_list]) 
                writer.writerow(["Attack time for each stego-file data set (s)", *directoryTimeLapseList])
                writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
                writer = csv.writer(output_file, delimiter=";")
                writer.writerow(["Active stego-attack type", attack_type])
                writer.writerow(["Total attack time lapse (s)", totalTimeLapseSec])
                writer.writerow(["Total attack time lapse (min)", totalTimeLapseMin])
                writer.writerow(["Stego-file data set", *data_set_type_list]) 
                writer.writerow(["Attack time for each stego-file data set (s)", *directoryTimeLapseList])
                writer.writerow('')

def log_time_attack_to_csv_all() -> None:
    base = "results/active_attacks_logs"
    all_attacks = "all_attacks.csv"
    with open(f"{base}/all_attacks.csv", "w", encoding="utf-8", newline="") as output_file:
        writer = csv.writer(output_file, delimiter=";")
        for attack_log_files in Path(base).iterdir():
            if attack_log_files.is_file() and not attack_log_files.name.startswith(".") and attack_log_files.name != all_attacks:
                with open(attack_log_files, "r", encoding="utf-8") as input_file:
                    reader = csv.reader(input_file, delimiter=";")
                    for row in reader:
                        writer.writerow(row)
                writer.writerow('')  # Add an empty line between different attack logs

def clean_logs_individual(attack_type: str) -> None:
    file = Path(f"results/active_attacks_logs/{attack_type}.csv")
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)

def clean_logs_all_attacks() -> None:
    file = Path("results/active_attacks_logs/all_attacks.csv")
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)

def clean_logs_folder() -> None:
    base = "results/active_attacks_logs"
    for attack_log_files in Path(base).iterdir():
        if attack_log_files.is_file() and not attack_log_files.name.startswith("."):
            os.remove(attack_log_files)

def unified_attack(
        attack_type: str, 
        attack: Callable[[str], DocumentObject]
        ) -> None:
    base = "data_set/stego_files"
    attacked_base = "data_set/attacked_stego_files"
    totalTimeLapse = 0
    directoryTimeLapseList = []
    data_set_type_list = []

    for directories in Path(base).iterdir():
        directoryTimeLapse = 0
        for file in directories.iterdir():
            print()
            if file.name.startswith('~$') or file.name.startswith('.'):
                continue
            start = time.time()

            docPath = str(Path(f"{base}/{directories.name}/{file.name}")) #Path("data_set/clean_files/TEST_0.docx")
            print(f"Opened: {docPath}")
            document = attack(docPath)

            stegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}"))
            document.save(stegoDocPath)
            print(f"Saved: {stegoDocPath}")

            end = time.time()
            fileTimeLapse = end - start
            directoryTimeLapse += fileTimeLapse
            print(f"Time taken (s): {int(fileTimeLapse * 100) / 100}")
            print()
        directoryTimeLapseSec = int(directoryTimeLapse * 100) / 100

        directoryTimeLapseTotalSec = int(directoryTimeLapse)    # truncate, no rounding
        directoryTimeLapsePureMin = directoryTimeLapseTotalSec // 60
        directoryTimeLapseSecRemainder = directoryTimeLapseTotalSec % 60

        directoryTimeLapseFloat = float(f"{directoryTimeLapsePureMin}.{directoryTimeLapseSecRemainder:02d}")

        data_set_type_list.append(directories.name)
        directoryTimeLapseList.append(directoryTimeLapseFloat)
        totalTimeLapse += directoryTimeLapse

        print(f"Directory timelapse (s): {directoryTimeLapseSec}")
        print(f"Directory timelapse (min): {directoryTimeLapsePureMin}")
        print(f"Directory timelapse: {directoryTimeLapseFloat}")
    print()
    
    totalTimeLapseTotalSec = int(totalTimeLapse)
    totalTimeLapseSec = int(totalTimeLapse * 100) / 100
    totalTimeLapsePureMin = totalTimeLapseTotalSec // 60
    totalTimeLapseSecRemainder = totalTimeLapseTotalSec % 60

    totalTimeLapseFloat = float(f"{totalTimeLapsePureMin}.{totalTimeLapseSecRemainder:02d}")

    print(f"Total timelapse (s): {totalTimeLapseSec}")
    print(f"Total timelapse (min): {totalTimeLapsePureMin}")
    print(f"Total timelapse: {totalTimeLapseFloat}")

    clean_logs_individual(attack_type)
    log_time_attack_to_csv_individual(attack_type, totalTimeLapseSec, totalTimeLapseFloat, data_set_type_list, directoryTimeLapseList)

if __name__ == "__main__":
    #clean_logs_folder()
    clean_logs_all_attacks()
    log_time_attack_to_csv_all()