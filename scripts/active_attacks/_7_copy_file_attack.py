import shutil
from pathlib import Path

base = "data_set/stego_files"
for directories in Path(base).iterdir():
    for file in directories.iterdir():
        print()
        if file.name.startswith('~$'):
            continue
        docPath = f"{base}/{directories.name}/{file.name}"
        print("Opened:", docPath)
        stegoDocPath = Path(f"data_set/attacked_stego_files/7_copy_attack/{directories.name}/{file.name}")
        print("Attacking:", stegoDocPath)
        shutil.copy(docPath, stegoDocPath)
        print("Saved:", stegoDocPath)