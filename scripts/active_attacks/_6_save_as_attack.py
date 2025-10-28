import os
import win32com.client as win32
from pathlib import Path
import time

base = "data_set/stego_files"
stego_dir = "data_set/attacked_stego_files/6_save_as_attack"
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
        # Ignore temporary files
        if file.name.startswith('~$'):
            continue
        start = time.time()
        docPath = Path(f"{base}/{directories.name}/{file.name}").resolve()
        print("Opened:", docPath)

        stegoDocPath = Path(f"{stego_dir}/{directories.name}/{file.name}").resolve()
        print("Attacking:", stegoDocPath)
        document = word.Documents.Open(str(docPath), ReadOnly=1, AddToRecentFiles=False)
        document.SaveAs2(str(stegoDocPath), FileFormat=wdFormatDocumentDefault)
        document.Close()

        print("Saved:", stegoDocPath)
        end = time.time()
        fileTimeLapse = end - start
        directoryTimeLapse += fileTimeLapse
        print("Time taken (s):", fileTimeLapse)
        print()
    totalTimeLapse += directoryTimeLapse
    print("Directory timelapse (s):", directoryTimeLapse)
    print("Directory timelapse (min):", directoryTimeLapse / 60)
word.Quit()

print()
print("Total timelapse (s):", totalTimeLapse)
print("Total timelapse (min):", totalTimeLapse / 60)