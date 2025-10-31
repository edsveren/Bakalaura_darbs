from pathlib import Path
from copy import deepcopy
import time
from docx import Document
from docx.oxml.parser import OxmlElement
from docx.document import Document as DocumentObject
from logging_time import *

def retype_in_new_document(
        document: DocumentObject, 
        new_document: DocumentObject
        ) -> None:
    for paragraph in document.paragraphs:
        new_paragraph_element = OxmlElement('w:p')
        for run in paragraph.runs:
            current_run_element = run._r
            new_run_element = OxmlElement('w:r')
            for child_element in current_run_element:
                new_run_element.append(deepcopy(child_element))
            new_paragraph_element.append(new_run_element)
        new_document._element.body.append(new_paragraph_element)

if __name__ == "__main__":
    base = "data_set/stego_files"
    attacked_base = "data_set/attacked_stego_files"
    attack_type = "8_retype_attack"
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
            new_document = Document()
            retype_in_new_document(document, new_document)
            stegoDocPath = str(Path(f"{attacked_base}/{attack_type}/{directories.name}/{file.name}"))
            new_document.save(str(stegoDocPath))
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