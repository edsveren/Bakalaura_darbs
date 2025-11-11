import os
import csv
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject

def count_total_paragraphs(document: DocumentObject) -> int:
    count = 0
    for _ in document.paragraphs:
        count += 1
    return count

def count_total_runs_elements(document: DocumentObject) -> int:
    count = 0
    for paragraph in document.paragraphs:
        for _ in paragraph.runs:
            count += 1
    return count

def count_total_text_elements(document: DocumentObject) -> int:
    text_element_count = 0
    text_element = f".//{qn('w:t')}"
    #root = document.part.element
    #docTextElement = root.xpath("//w:document/w:body/w:p/w:r/w:t")
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            text_element_in_run_count = len(run._r.findall(text_element)) # ElementPath findall() function - returns a list of matching Elements
            text_element_count += text_element_in_run_count
    return text_element_count

def count_text_elements_per_paragraph(document: DocumentObject) -> list[int]:
    text_element_per_paragraph = []
    text_element = f".//{qn('w:t')}"
    for paragraph in document.paragraphs:
        text_element_in_run_count = len(paragraph._p.findall(text_element))
        text_element_per_paragraph.append(text_element_in_run_count)
    return text_element_per_paragraph

def count_text_elements_per_run(document: DocumentObject) -> list[int]:
    text_element_per_run = []
    text_element = f".//{qn('w:t')}"
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            text_element_in_run_count = len(run._r.findall(text_element))
            text_element_per_run.append(text_element_in_run_count)
    return text_element_per_run

def to_csv(docPath: Path, data_set: str, total_paragraph_count: int|None, 
           total_run_element_count: int|None, total_text_elements: int|None, 
           text_element_per_paragraph_list: list[int]|None, text_element_per_run_list: list[int]|None) -> None:
    file_name = docPath.stem
    result_file = f"results/statistical_analysis/1_element_count/{file_name}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            if total_paragraph_count != None:
                writer.writerow(["Total Paragraph Count", total_paragraph_count])
            if total_run_element_count != None:
                writer.writerow(["Total Run Element Count", total_run_element_count])
            if total_text_elements != None:
                writer.writerow(["Total Text Element Count", total_text_elements])
            if text_element_per_paragraph_list != None:
                writer.writerow(["Text Elements Per Paragraph", *text_element_per_paragraph_list])  # *text_element_per_paragraph_list 
            if text_element_per_run_list != None:
                writer.writerow(["Text Elements Per Run", *text_element_per_run_list])  # *text_element_per_run_list
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            #w.writerow([file_name, text_element_count, *font_sizes])
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            if total_paragraph_count != None:
                writer.writerow(["Total Paragraph Count", total_paragraph_count])
            if total_run_element_count != None:
                writer.writerow(["Total Run Element Count", total_run_element_count])
            if total_text_elements != None:
                writer.writerow(["Total Text Element Count", total_text_elements])
            if text_element_per_paragraph_list != None:
                writer.writerow(["Text Elements Per Paragraph", *text_element_per_paragraph_list])  # *text_element_per_paragraph_list 
            if text_element_per_run_list != None:
                writer.writerow(["Text Elements Per Run", *text_element_per_run_list])  # *text_element_per_run_list
            writer.writerow('')

def main() -> None:
    if Path(f"results/statistical_analysis/1_element_count/TEST_0.csv").is_file():
        os.remove(Path(f"results/statistical_analysis/1_element_count/TEST_0.csv"))
    docPath_0 = Path("data_set/clean_files/TEST_0.docx")
    docPath_1 = Path("data_set/stego_files/stego_method_1/TEST_0.docx")
    docPath_2 = Path("data_set/stego_files/stego_method_2/TEST_0.docx")
    docPath_3 = Path("data_set/stego_files/stego_method_3/TEST_0.docx")
    docPath_4 = Path("data_set/stego_files/stego_method_4/TEST_0.docx")
    docPath_5 = Path("data_set/stego_files/stego_method_5/TEST_0.docx")
    docPath_6= Path("data_set/stego_files/stego_method_6/TEST_0.docx")

    paths = [docPath_0, docPath_1, docPath_2, docPath_3, docPath_4, docPath_5, docPath_6]
    data_set = ["clean", "hide_in_text", "multilayer_hybrid", "two_bit_transformation", "modify_RGB_color_ch", "unispace", "unicode_homoglyphs"]

    i = 0
    for path in paths:
        print("")
        if not Path(path).is_file():
            print(f"File doesn't exist: {path}")
            continue
        print(f"Opened: {path}")
        document = Document(str(path))
        total_paragraph_count = count_total_paragraphs(document)
        total_run_element_count = count_total_runs_elements(document)
        total_text_elements = count_total_text_elements(document)
        text_element_per_paragraph_list = count_text_elements_per_paragraph(document)
        text_element_per_run_list = count_text_elements_per_run(document)

        print(f"Total paragraph count: {total_paragraph_count}")
        print(f"Total run element count: {total_run_element_count}")
        print(f"Total text element count: {total_text_elements}")
        print(f"Text elements per paragraph: {text_element_per_paragraph_list}")
        print(f"Text elements per run: {text_element_per_run_list}")
        # print(get_font_size_value_from_each_run(document, i, 0, path, "clean"))
        to_csv(path, data_set[i], total_paragraph_count, total_run_element_count, total_text_elements, text_element_per_paragraph_list, None)
        i += 1

if __name__ == "__main__":
    main()