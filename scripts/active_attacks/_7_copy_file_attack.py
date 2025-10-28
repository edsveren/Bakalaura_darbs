import shutil
from pathlib import Path

base = "data_set/stego-files"
for directories in Path(base).iterdir():
    for file in directories.iterdir():
        docPath = f"{base}/{directories.name}/{file.name}"
        stegoDocPath = Path(f"data_set/attacked_stego-files/7_copy_attack/{directories.name}/{file.name}")
        shutil.copy(docPath, stegoDocPath)
        print("Saved:", stegoDocPath)
