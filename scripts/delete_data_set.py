import os
from pathlib import Path

def delete_stego_data_set():
    stego_files = "data_set/stego_files"
    for directories in Path(stego_files).iterdir():
        for file in directories.iterdir():
            #docPath = f"{stego_files}/{directories.name}/{file.name}"
            if file.is_file() and not file.name.startswith("."):
                os.remove(file)
                print(f"Deleted: {file} in {directories.name}")

def delete_attacked_data_set():
    attacked_stego_files = "data_set/attacked_stego_files"
    #folder_ensurer = ".gitkeep"
    for attack_directories in Path(attacked_stego_files).iterdir():
        for stego_directories in attack_directories.iterdir():
            for file in stego_directories.iterdir():
                #docPath = f"{attacked_stego_files}/{stego_directories.name}/{file.name}"
                if file.is_file() and not file.name.startswith("."):
                    os.remove(file)
                    print(f"Deleted: {file} in {stego_directories}")

if __name__ == "__main__":
    delete_stego_data_set()
    delete_attacked_data_set()