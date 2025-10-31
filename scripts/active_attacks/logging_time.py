import os
import csv
from pathlib import Path

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

if __name__ == "__main__":
    #clean_logs_folder()
    clean_logs_all_attacks()
    log_time_attack_to_csv_all()