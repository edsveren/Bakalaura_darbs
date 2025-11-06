import shutil
from pathlib import Path
import time
from logging_time import *

if __name__ == "__main__":
    base = "data_set/stego_files"
    attacked_base = "data_set/attacked_stego_files"
    attack_type = "07_copy_attack"
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
            docPath = f"{base}/{directories.name}/{file.name}"
            print(f"Opened: {docPath}")
            
            print(f"Attacking: {docPath}")
            stegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}"))
            shutil.copy(docPath, stegoDocPath)
            print(f"Saved: {stegoDocPath}")
            end = time.time()
            fileTimeLapse = end - start
            directoryTimeLapse += fileTimeLapse
            print(f"Time taken (s): {round(fileTimeLapse, 2)}")
            print()
        data_set_type_list.append(directories.name)
        directoryTimeLapseList.append(round(directoryTimeLapse, 2))
        totalTimeLapse += directoryTimeLapse
        print(f"Directory timelapse (s): {round(directoryTimeLapse, 2)}")
        print(f"Directory timelapse (min): {round(directoryTimeLapse / 60, 2)}")
    print()
    totalTimeLapseSec = round(totalTimeLapse, 2)
    totalTimeLapseMin = round(totalTimeLapse / 60, 2)
    print(f"Total timelapse (s): {totalTimeLapseSec}")
    print(f"Total timelapse (min): {totalTimeLapseMin}")

    clean_logs_individual(attack_type)
    log_time_attack_to_csv_individual(attack_type, totalTimeLapseSec, totalTimeLapseMin, data_set_type_list, directoryTimeLapseList)