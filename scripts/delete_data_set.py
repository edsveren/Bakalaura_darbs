from pathlib import Path
import os

def delete_stego_data_set():
    stego_files = "data_set/stego_files"
    for directories in Path(stego_files).iterdir():
        for file in directories.iterdir():
            docPath = f"{stego_files}/{directories.name}/{file.name}"
            if Path(docPath).is_file():
                os.remove(Path(docPath))
                print(f"Deleted: {docPath} in {directories.name}")

def delete_attacked_data_set():
    attacked_stego_files = "data_set/attacked_stego_files"
    for attack_directories in Path(attacked_stego_files).iterdir():
        for stego_directories in attack_directories.iterdir():
            for file in stego_directories.iterdir():
                docPath = f"{attacked_stego_files}/{stego_directories.name}/{file.name}"
                if Path(docPath).is_file():
                    os.remove(Path(docPath))
                    print(f"Deleted: {docPath} in {stego_directories.name}")

if __name__ == "__main__":
    delete_stego_data_set()
    #delete_attacked_data_set()