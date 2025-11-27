import os
import csv
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
from typing import Callable

# Delete file
def delete_file(file: Path) -> None:
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)
        print(f"Deleted: {str(file)}")

def export_to_csv(
        docPath: Path, 
        data_to_csv: list[list]
    ) -> None:
    file_name = docPath.stem
    result_file = f"results/statistical_analysis/1_element_count/{file_name}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            for row in data_to_csv:
                writer.writerow(row)
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            for row in data_to_csv:
                writer.writerow(row)
            writer.writerow('')

def singular_check(
        statistical_analysis_method: str, 
        analyzed_file: str, 
        statistical_analysis_method_execution: Callable[[Path, str], list[list]]
    ) -> None:

    csv_file = f"results/statistical_analysis/{statistical_analysis_method}/{analyzed_file}.csv"
    delete_file(Path(csv_file))

    paths = []
    paths.append(Path(f"data_set/clean_files/{analyzed_file}.docx"))
    for i in range(1, 7):
        docPath = Path(f"data_set/stego_files/stego_method_{i}/{analyzed_file}.docx")
        paths.append(docPath)
    # paths = [docPath_0, docPath_1, docPath_2, docPath_3, docPath_4, docPath_5, docPath_6]

    data_set = ["clean", "hide_in_text", "multilayer_hybrid", "two_bit_transformation", "modify_RGB_color_ch", "unispace", "unicode_homoglyphs"]

    j = 0
    for path in paths:
        print("")
        if not Path(path).is_file():
            print(f"File doesn't exist: {path}")
            j += 1
            continue
        print(f"Opened: {path}")

        data_to_csv = statistical_analysis_method_execution(path, data_set[j])

        # for item in data_to_csv:
        #     print(f"{item[0]}: {item[1]}")

        # Temporary solution
        for item in data_to_csv:
            label, *values = item
            print(f"{label}: {', '.join(map(str, values))}")

        export_to_csv(path, data_to_csv)
        j += 1