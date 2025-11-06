import re
from pathlib import Path
import time
from docx import Document
from docx.text.run import Run
from docx.oxml.ns import qn
from logging_time import *

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

if __name__ == "__main__":
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
            docPath = f"{base}/{directories.name}/{file.name}" #Path("data_set/clean_files/TEST_0.docx")
            print(f"Opened: {docPath}")
            
            print(f"Attacking: {docPath}")
            document = Document(docPath)
            every_nth_word = 10
            word_index = 0
            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    run_element = run._r
                    if run_element.find(qn('w:t')) != None:
                        word_index = delete_every_nth_word(run, every_nth_word, word_index)

            stegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}"))
            document.save(stegoDocPath)
            print(f"Saved: {stegoDocPath}")
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
    print()
    totalTimeLapseSec = round(totalTimeLapse, 2)
    totalTimeLapseMin = round(totalTimeLapse / 60, 2)
    print(f"Total timelapse (s): {totalTimeLapseSec}")
    print(f"Total timelapse (min): {totalTimeLapseMin}")

    clean_logs_individual(attack_type)
    log_time_attack_to_csv_individual(attack_type, totalTimeLapseSec, totalTimeLapseMin, data_set_type_list, directoryTimeLapseList)