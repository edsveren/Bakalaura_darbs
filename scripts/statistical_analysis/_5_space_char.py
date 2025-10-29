import re
import os
import csv
from pathlib import Path
from docx import Document
from docx.document import Document as DocumentObject

def count_words_in_paragraphs(document: DocumentObject) -> int:
    word_count = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        words = re.findall(r'\S+', text, flags=re.UNICODE)
        word_count += len(words)
    return word_count

def count_whitespace_characters_in_each_paragraph(document: DocumentObject) -> list:
    whitespace_counts = []
    for paragraph in document.paragraphs:
        whitespace_count = sum(1 for char in paragraph.text if char.isspace())
        whitespace_counts.append(whitespace_count)
    with open("results/whitespace_counts.txt", "w", encoding="utf-8") as file:
        for count in whitespace_counts:
            file.write(str(count) + "\n")
    return whitespace_counts

def count_whitespace_characters_in_each_run(document: DocumentObject) -> list:
    whitespace_counts = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            whitespace_count = sum(1 for char in run.text if char.isspace())
            whitespace_counts.append(whitespace_count)
    with open("results/whitespace_counts.txt", "w", encoding="utf-8") as file:
        for count in whitespace_counts:
            file.write(str(count) + "\n")
    return whitespace_counts

def to_csv(docPath: Path, data_set: str, total_word_count: int, total_whitespace_count: int, 
           word_to_whitespace_ratio: float, whitespace_counts_per_paragraph: list) -> None:
    file_name = docPath.stem
    result_file = f"results/5_space_char/{file_name}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Total word count", total_word_count])
            writer.writerow(["Total whitespace count", total_whitespace_count])
            writer.writerow(["Word to Whitespace ratio (%)", word_to_whitespace_ratio])
            writer.writerow(["Whitespace counts per paragraph", *whitespace_counts_per_paragraph])
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Total word count", total_word_count])
            writer.writerow(["Total whitespace count", total_whitespace_count])
            writer.writerow(["Word to Whitespace ratio (%)", word_to_whitespace_ratio])
            writer.writerow(["Whitespace counts per paragraph", *whitespace_counts_per_paragraph])
            writer.writerow('')

if __name__ == "__main__":
    if Path(f"results/5_space_char/TEST_0.csv").is_file():
        os.remove(Path(f"results/5_space_char/TEST_0.csv"))

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
        total_word_count = count_words_in_paragraphs(document)
        whitespace_counts_per_paragraph = count_whitespace_characters_in_each_paragraph(document)
        whitespace_counts_per_run = count_whitespace_characters_in_each_run(document)
        total_whitespace_count = 0
        for count in whitespace_counts_per_run:
            total_whitespace_count += count
        word_to_whitespace_ratio = round((total_whitespace_count/total_word_count) * 100, 2)
        
        print(f"Whitespace count in each paragraph: {whitespace_counts_per_paragraph}")
        print(f"Whitespace count in each run: {whitespace_counts_per_run}")
        # print(f"Total whitespace count: {total_whitespace_count}")
        # print(f"Total word count: {total_word_count}")
        print(f"Total word count: {total_word_count} to total whitespace count {total_whitespace_count}. Ratio: {word_to_whitespace_ratio} (%)")
        to_csv(path, data_set[i], total_word_count, total_whitespace_count, word_to_whitespace_ratio, whitespace_counts_per_paragraph)
        i += 1
