from docx import Document
from pathlib import Path
from collections import Counter
import _1_element_count
import os
import csv

def number_of_runs_with_a_single_character(document: Document) -> int:
    count = 0
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if len(run.text) == 1 and run.text != ' ':
                count += 1
    return count

def run_text_with_single_character(document: Document) -> list:
    runs = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            if len(run.text) == 1 and run.text != ' ':
                runs.append(run.text)
    with open("results/one_char.txt", "w", encoding="utf-8") as file:
        for char in runs:
            file.write(str(char) + "\n")
    return runs

def bin_single_chars(single_char: str) -> str:
    if single_char.isupper():   
        return "uppercase"
    elif single_char.islower():
        return "lowercase"
    elif single_char.isdigit():
        return "digit"
    else:
        return "other"
    
def to_csv(docPath: Path, data_set: str, chars: list, char_frequencies: list, 
           frequency_percentages_single_run: list, frequency_percentages_total_runs: list) -> None:
    file_name = docPath.stem
    result_file = f"results/4_single_char/{file_name}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Chars", *chars])
            writer.writerow(["Char frequencies", *char_frequencies])
            writer.writerow(["Char frequencies (single run, %)", *frequency_percentages_single_run])
            writer.writerow(["Char frequencies (total runs, %)", *frequency_percentages_total_runs])
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["Chars", *chars])
            writer.writerow(["Char frequencies", *char_frequencies])
            writer.writerow(["Char frequencies (single run, %)", *frequency_percentages_single_run])
            writer.writerow(["Char frequencies (total runs, %)", *frequency_percentages_total_runs])
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
    if Path(f"results/4_single_char/TEST_0.csv").is_file():
        os.remove(Path(f"results/4_single_char/TEST_0.csv"))

    i = 0
    for path in paths:
        print("")
        print(f"Opened: {path}")
        document = Document(path)
        total_run_element_count = _1_element_count.count_total_runs_elements(document)
        single_char_run_count = number_of_runs_with_a_single_character(document)
        run_percentages = str(round((single_char_run_count / total_run_element_count) * 100, 2)).replace(".", ",")
        single_char_texts = run_text_with_single_character(document)
        single_char_texts_list = [bin_single_chars(single_char) for single_char in single_char_texts]
        single_char_bins = {key: Counter(single_char_texts_list).get(key, 0) for key in ['uppercase', 'lowercase', 'digit', 'other']}
        
        chars = []
        frequencies = []
        frequency_percentages_single_run = []
        frequency_percentages_total_runs = []

        print(f"Number of runs with a single char (non-whitespace): {single_char_run_count} out of {total_run_element_count} runs. ({run_percentages}%).")
        # print("Char from each single-char run element:")
        # print(single_char_texts)
        for char, frequency in single_char_bins.items():
            frequency_percent_out_of_single_run = str(round((frequency / single_char_run_count) * 100, 2)) #.replace(".", ",")
            frequency_percent_out_of_total_runs = str(round((frequency / total_run_element_count) * 100, 2)) #.replace(".", ",")
            chars.append(char)
            frequencies.append(frequency)
            frequency_percentages_single_run.append(frequency_percent_out_of_single_run)
            frequency_percentages_total_runs.append(frequency_percent_out_of_total_runs)
            print(f"Char type: {char}. Frequency: {frequency} (Single char runs: {frequency_percent_out_of_single_run}%, Total runs: {frequency_percent_out_of_total_runs}%).")
        to_csv(path, data_set[i], chars, frequencies, frequency_percentages_single_run, frequency_percentages_total_runs)
        i += 1