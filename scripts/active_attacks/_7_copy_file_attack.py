import shutil
from pathlib import Path
import time
import logging_time

def copy_attack(
        docPath: str, 
        stegoDocPath: str
        ) -> None:
    print(f"Attacking: {docPath}")
    shutil.copy(docPath, stegoDocPath)

# def main() -> None:
#     base = "data_set/stego_files"
#     attacked_base = "data_set/attacked_stego_files"
#     attack_type = "07_copy_attack"
#     totalTimeLapse = 0
#     directoryTimeLapseList = []
#     data_set_type_list = []
#     for directories in Path(base).iterdir():
#         directoryTimeLapse = 0
#         for file in directories.iterdir():
#             print()
#             if file.name.startswith('~$') or file.name.startswith('.'):
#                 continue
#             start = time.time()

#             docPath = str(Path(f"{base}/{directories.name}/{file.name}"))
#             stegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}"))

#             print(f"Opened: {docPath}")
#             copy_attack(docPath, stegoDocPath)
#             print(f"Saved: {stegoDocPath}")

#             end = time.time()
#             fileTimeLapse = end - start
#             directoryTimeLapse += fileTimeLapse
#             print(f"Time taken (s): {int(fileTimeLapse * 100) / 100}")
#             print()
#         directoryTimeLapseSec = int(directoryTimeLapse * 100) / 100

#         directoryTimeLapseTotalSec = int(directoryTimeLapse)    # truncate, no rounding
#         directoryTimeLapsePureMin = directoryTimeLapseTotalSec // 60
#         directoryTimeLapseSecRemainder = directoryTimeLapseTotalSec % 60

#         directoryTimeLapseFloat = float(f"{directoryTimeLapsePureMin}.{directoryTimeLapseSecRemainder:02d}")

#         data_set_type_list.append(directories.name)
#         directoryTimeLapseList.append(directoryTimeLapseFloat)
#         totalTimeLapse += directoryTimeLapse

#         print(f"Directory timelapse (s): {directoryTimeLapseSec}")
#         print(f"Directory timelapse (min): {directoryTimeLapsePureMin}")
#         print(f"Directory timelapse: {directoryTimeLapseFloat}")
#     print()
    
#     totalTimeLapseTotalSec = int(totalTimeLapse)
#     totalTimeLapseSec = int(totalTimeLapse * 100) / 100
#     totalTimeLapsePureMin = totalTimeLapseTotalSec // 60
#     totalTimeLapseSecRemainder = totalTimeLapseTotalSec % 60

#     totalTimeLapseFloat = float(f"{totalTimeLapsePureMin}.{totalTimeLapseSecRemainder:02d}")

#     print(f"Total timelapse (s): {totalTimeLapseSec}")
#     print(f"Total timelapse (min): {totalTimeLapsePureMin}")
#     print(f"Total timelapse: {totalTimeLapseFloat}")

#     logging_time.clean_logs_individual(attack_type)
#     logging_time.log_time_attack_to_csv_individual(attack_type, totalTimeLapseSec, totalTimeLapseFloat, data_set_type_list, directoryTimeLapseList)

if __name__ == "__main__":
    # main()
    logging_time.unified_attack("07_copy_attack", copy_attack, False)