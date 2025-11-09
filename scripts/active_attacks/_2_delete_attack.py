import re
from pathlib import Path
import time
from docx import Document
from docx.text.run import Run
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import logging_time

def delete_every_nth_word(
        run: Run, 
        n: int, 
        word_index: int
        ) -> int:
    text = run.text
    text_pattern = re.compile(r'(\s+)')
    tokens = text_pattern.split(text)
    #print(tokens)
    #print("".join(tokens))
    text_list = []

    for token in tokens:
        if token.strip() == "":
            text_list.append(token)
            continue
        else:
            word_index += 1
            if word_index == (n - 1):
                text_list.append("")
                word_index = 0
            else:
                text_list.append(token)
    #print(text_list)
    run.text = "".join(text_list)
    #print(text)
    return word_index

def delete_attack(docPath: str) -> DocumentObject:
    print(f"Attacking: {docPath}")
    document = Document(docPath)
    every_nth_word = 10
    word_index = 0
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run_element = run._r
            if run_element.find(qn('w:t')) != None:
                word_index = delete_every_nth_word(run, every_nth_word, word_index)
    return document

def main() -> None:
    base = "data_set/stego_files"
    attacked_base = "data_set/attacked_stego_files"
    attack_type = "02_delete_attack"
    totalTimeLapse = 0
    directoryTimeLapseList = []
    data_set_type_list = []
    for directories in Path(base).iterdir():
        directoryTimeLapse = 0
        for file in directories.iterdir():
            print()
            if file.name.startswith('~$') or file.name.startswith('.'):
                continue
            start = time.time()

            docPath = str(Path(f"{base}/{directories.name}/{file.name}")) #Path("data_set/clean_files/TEST_0.docx")
            print(f"Opened: {docPath}")
            document = delete_attack(docPath)

            stegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}"))
            document.save(stegoDocPath)
            print(f"Saved: {stegoDocPath}")
            
            end = time.time()
            fileTimeLapse = end - start
            directoryTimeLapse += fileTimeLapse
            print(f"Time taken (s): {int(fileTimeLapse * 100) / 100}")
            print()
        directoryTimeLapseSec = int(directoryTimeLapse * 100) / 100

        directoryTimeLapseTotalSec = int(directoryTimeLapse)    # truncate, no rounding
        directoryTimeLapsePureMin = directoryTimeLapseTotalSec // 60
        directoryTimeLapseSecRemainder = directoryTimeLapseTotalSec % 60

        directoryTimeLapseFloat = float(f"{directoryTimeLapsePureMin}.{directoryTimeLapseSecRemainder:02d}")

        data_set_type_list.append(directories.name)
        directoryTimeLapseList.append(directoryTimeLapseFloat)
        totalTimeLapse += directoryTimeLapse

        print(f"Directory timelapse (s): {directoryTimeLapseSec}")
        print(f"Directory timelapse (min): {directoryTimeLapsePureMin}")
        print(f"Directory timelapse: {directoryTimeLapseFloat}")
    print()
    
    totalTimeLapseTotalSec = int(totalTimeLapse)
    totalTimeLapseSec = int(totalTimeLapse * 100) / 100
    totalTimeLapsePureMin = totalTimeLapseTotalSec // 60
    totalTimeLapseSecRemainder = totalTimeLapseTotalSec % 60

    totalTimeLapseFloat = float(f"{totalTimeLapsePureMin}.{totalTimeLapseSecRemainder:02d}")

    print(f"Total timelapse (s): {totalTimeLapseSec}")
    print(f"Total timelapse (min): {totalTimeLapsePureMin}")
    print(f"Total timelapse: {totalTimeLapseFloat}")

    logging_time.clean_logs_individual(attack_type)
    logging_time.log_time_attack_to_csv_individual(attack_type, totalTimeLapseSec, totalTimeLapseFloat, data_set_type_list, directoryTimeLapseList)

if __name__ == "__main__":
    main()