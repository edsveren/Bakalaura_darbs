import re
from pathlib import Path
from docx import Document
from docx.document import Document as DocumentObject
import unified_statistical_analysis_file

def count_chars_in_paragraphs(document: DocumentObject) -> int:
    char_count = 0
    for paragraph in document.paragraphs:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        chars = re.findall(r'[\s\S]', text, flags=re.UNICODE)
        char_count += len(chars)
    return char_count

def count_non_ascii_chars_in_paragraphs(document: DocumentObject) -> tuple[int, list[int]]:
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

def count_non_ascii_chars_in_runs(document: DocumentObject) -> list[int]:
    char_count_run = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run_text = run.text.replace('\xa0', '\x20')  # NBSP -> space
            run_chars = re.findall(r'[^\x00-\x7F]', run_text, flags=re.UNICODE)
            char_count_run.append(len(run_chars))
    return char_count_run

def unicode_analysis(path: Path, data_set: str) -> list[list]:
    document = Document(str(path))
    total_char_count = count_chars_in_paragraphs(document)
    total_non_ascii_char_count, non_ascii_char_count_paragraphs = count_non_ascii_chars_in_paragraphs(document)
    total_char_to_non_ascii_char_ratio = round((total_non_ascii_char_count/total_char_count) * 100, 2)

    # print(f"Non-ASCII characters in paragraphs: {non_ascii_char_count_paragraphs}")
    # # print(f"Total non-ASCII char count: {non_ascii_char_count}")
    # # print(f"Total char count: {char_count}")
    # print(f"Total char count: {total_char_count} to total non-ASCII char count: {total_non_ascii_char_count}. Ratio: {total_char_to_non_ascii_char_ratio} (%)")
    
    data_to_csv = [
        ["Document Name", path.stem],
        ["Data set", data_set],
        ["Total char count", total_char_count],
        ["Total ASCII char count", total_non_ascii_char_count],
        ["Total char to ASCII char ratio (%)", total_char_to_non_ascii_char_ratio],
        ["Non-ASCII counts per paragraph", *non_ascii_char_count_paragraphs]
    ]
    
    return data_to_csv
        
def main() -> None:
    unified_statistical_analysis_file.singular_check('6_unicode', 'TEST_0', unicode_analysis)

if __name__ == "__main__":
    main()