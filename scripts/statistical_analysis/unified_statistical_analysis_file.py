import os
import csv
from pathlib import Path
from typing import Callable

# Delete file
def delete_file(file: Path) -> None:
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)
        print(f"Deleted: {str(file)}")

def export_to_csv(
        result_file: str, 
        data_to_csv: list[list]
    ) -> None:

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

        export_to_csv(csv_file, data_to_csv)
        j += 1


# docPath_0 = Path("data_set/clean_files/TEST_0.docx")
# docPath_1 = Path("data_set/stego_files/stego_method_1/TEST_0.docx")
# docPath_3 = Path("data_set/stego_files/stego_method_3/TEST_0.docx")
# docPath_4 = Path("data_set/stego_files/stego_method_4/TEST_0.docx")
# docPath_5 = Path("data_set/stego_files/stego_method_5/TEST_0.docx")
# docPath_6= Path("data_set/stego_files/stego_method_6/TEST_0.docx")

# data_set = ["clean", "hide_in_text", "2_bit_transformation", "modify_RGB_color_ch", "unispace", "unicode_homoglyphs"]

# clean_files = "data_set/clean_files"
# i = 1
# for file in Path(clean_files).iterdir():
#     print()
#     if file.name.startswith('~$'):
#         continue
#     docPath = str(Path(f"{clean_files}/{file.name}")) #Path("data_set/clean_files/TEST_0.docx")
#     print(f"DOCUMENT: {docPath}")
#     document = Document(docPath)
#     txt_element_count = count_text_elements(document)
#     print(f"Text element count: {txt_element_count}")
#     print("Font sizes values for each run element:")
#     print(get_font_size_value_from_each_run(document, i, txt_element_count, docPath, data_set[0]))
#     i += 1

# stego_files = "data_set/stego_files"
# i = 0
# for directories in Path(stego_files).iterdir():
#     for file in directories.iterdir():
#         docPath = str(Path(f"{stego_files}/{directories.name}/{file.name}"))
#         print(f"DOCUMENT: {docPath}")
#         document = Document(docPath)
#         print(f"Text element count: {count_text_elements(document)}")
#         print("Font sizes values for each run element:")
#         print(get_font_size_value_from_each_run(document, docPath, data_set[i]))
#     i += 1