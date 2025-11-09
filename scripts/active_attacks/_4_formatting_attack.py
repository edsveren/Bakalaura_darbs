from pathlib import Path
import time
from typing import cast
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor
from docx.document import Document as DocumentObject
from docx.styles.style import CharacterStyle
import logging_time

def change_text_style_for_auto_text(document: DocumentObject) -> DocumentObject:
    styles = document.styles
    base_style = styles.add_style('stego_style', WD_STYLE_TYPE.CHARACTER)
    new_style = cast(CharacterStyle, base_style) # the same as CharacterStyle(base_style)

    new_style_font = new_style.font
    new_style_font.name = 'Britannic Bold'
    new_style_font.size = Pt(16)
    new_style_font.color.rgb = RGBColor(0x23, 0x23, 0x23)

    for paragraph in document.paragraphs:
        paragraph_element = paragraph._p
        paragraph_properties = paragraph_element.pPr
        if paragraph_properties != None:
            pStyle = paragraph_properties.pStyle
        else:
            pStyle = None
        for run in paragraph.runs:
            run_element = run._r
            run_properties = run_element.rPr
            if run_properties != None:
                rStyle = run_properties.rStyle
            else:
                rStyle = None
            if pStyle == None and rStyle == None:
                run.style = new_style
    return document

def formatting_attack(docPath: str) -> DocumentObject:
    print(f"Attacking: {docPath}")
    document = Document(docPath)
    document = change_text_style_for_auto_text(document)
    return document

def main() -> None:
    base = "data_set/stego_files"
    attacked_base = "data_set/attacked_stego_files"
    attack_type = "04_format_attack"
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
            document = formatting_attack(docPath)

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