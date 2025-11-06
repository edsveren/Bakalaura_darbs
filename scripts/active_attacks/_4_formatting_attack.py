from pathlib import Path
import time
from typing import cast
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor
from docx.document import Document as DocumentObject
from docx.styles.style import CharacterStyle
from logging_time import *

def change_text_style_for_auto_text(document: DocumentObject) -> None:
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


if __name__ == "__main__":
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
            docPath = f"{base}/{directories.name}/{file.name}" #Path("data_set/clean_files/TEST_0.docx")
            print(f"Opened: {docPath}")

            print(f"Attacking: {docPath}")
            document = Document(docPath)
            change_text_style_for_auto_text(document)
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