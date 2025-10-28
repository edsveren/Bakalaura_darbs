import os
import csv
from docx import Document
from pathlib import Path
from docx.oxml.ns import qn
from collections import Counter

def paragraph_default_font_size_value(styles_element) -> str | None:
    all_styles = styles_element.findall(f".//{qn('w:style')}")
    for style in all_styles:
        if style.get(qn('w:type')) == 'paragraph' and style.get(qn('w:default')) == '1':
            font_size = style.find(f".//{qn('w:sz')}'")
            if font_size != None:
                font_size_value = font_size.get(qn('w:val'))
                if font_size_value != None:
                    font_size_value = int(font_size_value) / 2.0
                    return font_size_value
    # If no default paragraph style found
    return None

def document_default_font_size_value(styles_element) -> str | None:
    run_properties_default_element = styles_element.find(f".//{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
    if run_properties_default_element != None:
        font_size = run_properties_default_element.find(qn('w:sz'))
        if font_size != None:
            font_size_value = font_size.get(qn('w:val'))
            if font_size_value != None:
                font_size_value = int(font_size_value) / 2.0
                return font_size_value
    else: # If no default document style found
        return None

def document_default_font_size_pt(document: Document) -> int:
    styles_element = document.styles.element

    font_size_value_pt = 11.0
    # paragraph_font_size_value_pt = styles_element.xpath(".//w:style[@w:type='paragraph' and @w:default='1']//w:rPr/w:sz/@w:val")
    # default_font_size_value_pt = styles_element.xpath(".//w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")

    paragraph_font_size_value_pt = paragraph_default_font_size_value(styles_element)
    default_font_size_value_pt = document_default_font_size_value(styles_element)

    if paragraph_font_size_value_pt:
        # font_size_value_pt = int(paragraph_font_size_value_pt[0]) / 2.0
        font_size_value_pt = paragraph_font_size_value_pt
        return font_size_value_pt
    elif default_font_size_value_pt:
        # font_size_value_pt = int(default_font_size_value_pt[0]) / 2.0
        font_size_value_pt = default_font_size_value_pt
        return font_size_value_pt
    else:
        return font_size_value_pt

def get_font_size_value_from_each_run(document: Document, index: int, text_element_count: int, docPath: str, data_set: str) -> list:
    default_pt = document_default_font_size_pt(document)
    font_sizes = []
    for paragraph in document.paragraphs:
        paragraph_font_style_size = paragraph.style.font.size
        for run in paragraph.runs:
            run_font_size = run.font.size
            run_font_style_size = run.style.font.size
            if run_font_size != None:
                font_sizes.append(run_font_size.pt)
            elif run_font_style_size != None:
                font_sizes.append(run_font_style_size.pt)
            elif paragraph_font_style_size != None:
                font_sizes.append(paragraph_font_style_size.pt)
            else:
                font_sizes.append(default_pt)
    return font_sizes

def to_csv(docPath: Path, data_set: str, font_sizes: list, font_size_frequency: list) -> None:
    file_name = docPath.stem
    result_file = f"results/2_font_sizes/{file_name}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Font Sizes (pt)", *font_sizes])
            writer.writerow(["Font Sizes Count", *font_size_frequency])
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Font Sizes (pt)", *font_sizes])
            writer.writerow(["Font Sizes Count", *font_size_frequency])
            writer.writerow('')

if __name__ == "__main__":
    if Path(f"results/2_font_sizes/TEST_0.csv").is_file():
        os.remove(Path(f"results/2_font_sizes/TEST_0.csv"))

    docPath_0 = Path("data_set/clean_files/TEST_0.docx")
    docPath_1 = Path("data_set/stego_files/stego_method_1/TEST_0.docx")
    docPath_3 = Path("data_set/stego_files/stego_method_3/TEST_0.docx")
    docPath_4 = Path("data_set/stego_files/stego_method_4/TEST_0.docx")
    docPath_5 = Path("data_set/stego_files/stego_method_5/TEST_0.docx")
    docPath_6= Path("data_set/stego_files/stego_method_6/TEST_0.docx")

    paths = [docPath_0, docPath_1, docPath_3, docPath_4, docPath_5, docPath_6]
    data_set = ["clean", "hide_in_text", "two_bit_transformation", "modify_RGB_color_ch", "unispace", "unicode_homoglyphs"]
    i = 0
    for path in paths:
        print("")
        print(f"Opened: {path}")
        document = Document(path)
        font_sizes = get_font_size_value_from_each_run(document, i, 0, path, "clean")
        counter = Counter(font_sizes)        
        sizes = list(counter.keys())
        frequencies = list(counter.values())

        print("Font sizes values for each run element:")
        print(font_sizes)
        print(f"Font sizes: {sizes}")
        print(f"Font size frequencies: {frequencies}")
        for size, frequency in counter.items():
            print(f"Font size: {size} pt. Frequency: {frequency}")
        to_csv(path, data_set[i], sizes, frequencies)
        i += 1

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
#     docPath = f"{clean_files}/{file.name}" #Path("data_set/clean_files/TEST_0.docx")
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
#         docPath = f"{stego_files}/{directories.name}/{file.name}"
#         print(f"DOCUMENT: {docPath}")
#         document = Document(docPath)
#         print(f"Text element count: {count_text_elements(document)}")
#         print("Font sizes values for each run element:")
#         print(get_font_size_value_from_each_run(document, docPath, data_set[i]))
#     i += 1