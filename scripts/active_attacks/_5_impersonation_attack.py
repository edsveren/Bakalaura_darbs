import os
from pathlib import Path
import time
import win32com.client as win32
from win32com.client.dynamic import CDispatch as dynamic_CDispatch
import logging_time

def impersonation_attack(
        word: dynamic_CDispatch, 
        docPath: str, 
        stegoPDFPath: str, 
        stegoDocPath: str,
        wdFormatPDF: int,
        wdFormatDocumentDefault: int
        ) -> None:
    # DOCX to PDF conversion
    print(f"Converting to: {stegoPDFPath}")
    document = word.Documents.Open(docPath, ReadOnly=1, AddToRecentFiles=False)
    document.SaveAs2(stegoPDFPath, FileFormat=wdFormatPDF)
    document.Close()

    # PDF back to DOCX conversion
    print(f"Converting to: {stegoDocPath}")
    document = word.Documents.Open(stegoPDFPath, ReadOnly=1, AddToRecentFiles=False)
    document.SaveAs2(stegoDocPath, FileFormat=wdFormatDocumentDefault)
    document.Close()

    # Remove any leftover PDF files
    if Path(stegoPDFPath).exists():
        os.remove(stegoPDFPath)

# def main() -> None:
#     base = "data_set/stego_files"
#     attacked_base = "data_set/attacked_stego_files"
#     attack_type = "05_impersonation_attack"
#     wdFormatPDF = 17
#     wdFormatDocumentDefault = 16
#     totalTimeLapse = 0
#     directoryTimeLapseList = []
#     data_set_type_list = []
#     word = win32.Dispatch("Word.Application")
#     word.Visible = False
#     word.DisplayAlerts = 0
#     # Close any open documents
#     while word.Documents.Count > 0:
#         word.Documents(1).Close(SaveChanges=0)
#     for directories in Path(base).resolve().iterdir():
#         directoryTimeLapse = 0
#         for file in directories.iterdir():
#             print()
#             if file.name.startswith('~$') or file.name.startswith('.'):
#                 continue
#             start = time.time()

#             docPath = str(Path(f"{base}/{directories.name}/{file.name}").resolve())
#             stegoPDFPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.stem}.pdf").resolve())
#             stegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}").resolve())

#             print(f"Opened: {docPath}")
#             impersonation_attack(word, docPath, stegoPDFPath, stegoDocPath, wdFormatPDF, wdFormatDocumentDefault)
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
    logging_time.unified_attack("05_impersonation_attack", impersonation_attack, True)