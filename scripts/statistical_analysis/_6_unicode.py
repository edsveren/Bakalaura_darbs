import re
import os
from docx import Document
from pathlib import Path
import csv

def count_chars_in_paragraphs(document: Document) -> int:
    char_count = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        chars = re.findall(r'[\s\S]', text, flags=re.UNICODE)
        char_count += len(chars)
    return char_count

def count_non_ascii_chars_in_paragraphs(document: Document) -> tuple[int, list[int]]:
    char_count_total = 0
    char_count_paragraph = []
    for paragraph in document.paragraphs:
        paragraph_text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        paragraph_chars = re.findall(r'[^\x00-\x7F]', paragraph_text, flags=re.UNICODE)
        paragraph_chars_count = len(paragraph_chars)
        char_count_paragraph.append(paragraph_chars_count)
        char_count_total += paragraph_chars_count
        # for run in paragraph.runs:
        #     run_text = run.text.replace('\xa0', '\x20')  # NBSP -> space
        #     run_chars = re.findall(r'[^\x00-\x7F]', run_text, flags=re.UNICODE)
        #     run_chars_count = len(run_chars)
        #     char_count_paragraph.append(run_chars_count)
        #     char_count_total += run_chars_count
    return char_count_total, char_count_paragraph

def count_non_ascii_chars_in_runs(document: Document) -> list[int]:
    char_count_run = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run_text = run.text.replace('\xa0', '\x20')  # NBSP -> space
            run_chars = re.findall(r'[^\x00-\x7F]', run_text, flags=re.UNICODE)
            char_count_run.append(len(run_chars))
    return char_count_run

def to_csv(docPath: Path, data_set: str, total_char_count: int, non_ascii_char_count: int, 
           total_char_to_non_ascii_char_ratio: float, non_ascii_char_count_paragraphs: list) -> None:
    file_name = docPath.stem
    result_file = f"results/6_unicode/{file_name}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Total char count", total_char_count])
            writer.writerow(["Total ASCII char count", non_ascii_char_count])
            writer.writerow(["Total char to ASCII char ratio (%)", total_char_to_non_ascii_char_ratio])
            writer.writerow(["Non-ASCII counts per paragraph", *non_ascii_char_count_paragraphs])
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Total char count", total_char_count])
            writer.writerow(["Total ASCII char count", non_ascii_char_count])
            writer.writerow(["Total char to ASCII char ratio (%)", total_char_to_non_ascii_char_ratio])
            writer.writerow(["Non-ASCII counts per paragraph", *non_ascii_char_count_paragraphs])
            writer.writerow('')

docPath_0 = Path("data_set/clean_files/TEST_0.docx")
docPath_1 = Path("data_set/stego_files/stego_method_1/TEST_0.docx")
docPath_3 = Path("data_set/stego_files/stego_method_3/TEST_0.docx")
docPath_4 = Path("data_set/stego_files/stego_method_4/TEST_0.docx")
docPath_5 = Path("data_set/stego_files/stego_method_5/TEST_0.docx")
docPath_6= Path("data_set/stego_files/stego_method_6/TEST_0.docx")

paths = [docPath_0, docPath_1, docPath_3, docPath_4, docPath_5, docPath_6]
data_set = ["clean", "hide_in_text", "two_bit_transformation", "modify_RGB_color_ch", "unispace", "unicode_homoglyphs"]

if __name__ == "__main__":
    if Path(f"results/6_unicode/TEST_0.csv").is_file():
        os.remove(Path(f"results/6_unicode/TEST_0.csv"))
    i = 0
    for path in paths:
        print("")
        print(f"Opened: {path}")
        document = Document(path)
        total_char_count = count_chars_in_paragraphs(document)
        total_non_ascii_char_count, non_ascii_char_count_paragraphs = count_non_ascii_chars_in_paragraphs(document)
        total_char_to_non_ascii_char_ratio = round((total_non_ascii_char_count/total_char_count) * 100, 2)

        print(f"Non-ASCII characters in paragraphs: {non_ascii_char_count_paragraphs}")
        # print(f"Total non-ASCII char count: {non_ascii_char_count}")
        # print(f"Total char count: {char_count}")
        print(f"Total char count: {total_char_count} to total non-ASCII char count: {total_non_ascii_char_count}. Ratio: {total_char_to_non_ascii_char_ratio} (%)")
        to_csv(path, data_set[i], total_char_count, total_non_ascii_char_count, total_char_to_non_ascii_char_ratio, non_ascii_char_count_paragraphs)
        i += 1
        