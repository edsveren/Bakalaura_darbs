import os
from pathlib import Path
import time
import win32com.client as win32
from logging_time import *

if __name__ == "__main__":
    base = "data_set/stego_files"
    attacked_base = "data_set/attacked_stego_files"
    #stego_dir = "data_set/attacked_stego_files/5_impersonation_attack"
    attack_type = "5_impersonation_attack"
    totalTimeLapse = 0
    directoryTimeLapseList = []
    data_set_type_list = []
    wdFormatPDF = 17
    wdFormatDocumentDefault = 16
    word = win32.Dispatch("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0
    totalTimeLapse = 0
    # Close any open documents
    while word.Documents.Count > 0:
        word.Documents(1).Close(SaveChanges=0)
    for directories in Path(base).resolve().iterdir():
        directoryTimeLapse = 0
        for file in directories.iterdir():
            print()
            if file.name.startswith('~$') or file.name.startswith('.'):
                continue
            start = time.time()
            docPath = Path(f"{base}/{directories.name}/{file.name}").resolve()
            print(f"Opened: {docPath}")

            stegoPDFPath = Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.stem}.pdf").resolve()
            stegoDocPath = Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}").resolve()

            # DOCX to PDF conversion
            print(f"Converting to: {str(stegoPDFPath)}")
            document = word.Documents.Open(str(docPath), ReadOnly=1, AddToRecentFiles=False)
            document.SaveAs2(str(stegoPDFPath), FileFormat=wdFormatPDF)
            document.Close()

            # PDF back to DOCX conversion
            print(f"Converting to: {str(stegoDocPath)}")
            document = word.Documents.Open(str(stegoPDFPath), ReadOnly=1, AddToRecentFiles=False)
            document.SaveAs2(str(stegoDocPath), FileFormat=wdFormatDocumentDefault)
            document.Close()

            # Remove any leftover PDF files
            if stegoPDFPath.exists():
                os.remove(stegoPDFPath)

            print(f"Saved: {str(stegoDocPath)}")
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
    word.Quit()
    print()
    totalTimeLapseSec = round(totalTimeLapse, 2)
    totalTimeLapseMin = round(totalTimeLapse / 60, 2)
    print(f"Total timelapse (s): {totalTimeLapseSec}")
    print(f"Total timelapse (min): {totalTimeLapseMin}")

    clean_logs_individual(attack_type)
    log_time_attack_to_csv_individual(attack_type, totalTimeLapseSec, totalTimeLapseMin, data_set_type_list, directoryTimeLapseList)