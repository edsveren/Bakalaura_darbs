import win32com.client as win32
from pathlib import Path

base = "data_set/stego_files"
stego_dir = "data_set/attacked_stego_files/9_clear_format_attack"
wdFormatDocumentDefault = 16
word = win32.Dispatch("Word.Application")
word.Visible = False
word.DisplayAlerts = 0
# Close any open documents
while word.Documents.Count > 0:
    word.Documents(1).Close(SaveChanges=0)
for directories in Path(base).resolve().iterdir():
    directoryTimeLapse = 0
    for file in directories.iterdir():
        # Ignore temporary files
        if file.name.startswith('~$'):
            continue
        docPath = Path(f"{base}/{directories.name}/{file.name}").resolve()
        print("Opened:", docPath)

        stegoDocPath = Path(f"{stego_dir}/{directories.name}/{file.name}").resolve()
        print("Attacking:", stegoDocPath)
        document = word.Documents.Open(str(docPath), ReadOnly=1, AddToRecentFiles=False)
        # CTRL + A
        #document.Content.Select()
        word.Selection.WholeStory()
        # Click Clear All Formatting button
        #word.CommandBars.ExecuteMso("ClearFormatting")
        word.Selection.ClearFormatting()
        # Save
        document.SaveAs2(str(stegoDocPath), FileFormat=wdFormatDocumentDefault)
        document.Close()

        print("Saved:", stegoDocPath)
        print()
word.Quit()