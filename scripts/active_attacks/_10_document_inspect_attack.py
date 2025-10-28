import win32com.client as win32
from pathlib import Path

base = "data_set/stego_files"
stego_dir = "data_set/attacked_stego_files/10_document_inspector_attack"
wdFormatDocumentDefault = 16
wdRDIAll = 99
word = win32.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
# Close any open documents
while word.Documents.Count > 0:
    word.Documents(1).Close(SaveChanges=0)
for directories in Path(base).resolve().iterdir():
    for file in directories.iterdir():
        # Ignore temporary files
        if file.name.startswith('~$'):
            continue
        docPath = Path(f"{base}/{directories.name}/{file.name}").resolve()
        print("Opened:", docPath)

        stegoDocPath = Path(f"{stego_dir}/{directories.name}/{file.name}").resolve()
        print("Attacking:", stegoDocPath)

        # Document Inspector Attack for Metadata
        document = word.Documents.Open(str(docPath), ReadOnly=1, AddToRecentFiles=False)

        # Document Inspector Attack for Content
        document.RemoveDocumentInformation(RemoveDocInfoType=wdRDIAll)
        for i in range(1, document.DocumentInspectors.Count + 1):
            status, result = document.DocumentInspectors.Item(i).Inspect()
            print(i, document.DocumentInspectors.Item(i).Name, "Status:", status, "Result:", result)
            if status == 1: 
                document.DocumentInspectors.Item(i).Fix()

        # Save
        document.SaveAs2(str(stegoDocPath), FileFormat=wdFormatDocumentDefault)
        document.Close()
        print("Saved:", stegoDocPath)
        print()
word.Quit()