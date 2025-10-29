import os
import csv
from pathlib import Path
from collections import Counter
from docx import Document
from docx.document import Document as DocumentObject

def get_RGB_value_from_each_run(document: DocumentObject) -> list:
    rgb_values = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color
            if font_color.rgb != None:
                rgb_values.append(f"#{font_color.rgb}")
            else:
                rgb_values.append("#000000")
    # with open("results/rgb_values.txt", "w", encoding="utf-8") as file:
    #     for rgb in rgb_values:
    #         file.write(str(rgb) + "\n")
    return rgb_values

def bin_rgb_values(rgb_value: str) -> str:
    match rgb_value:
        case "#000000":   
            return "black"
        case "#FFFFFF":
            return "white"
        case _:
            return "other"

def to_csv(docPath: Path, data_set: str, rgb_colours: list, rgb_frequencies: list, rgb_frequency_percentages: list) -> None:
    file_name = docPath.stem
    result_file = f"results/3_RGB_value/{file_name}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["RGB colours", *rgb_colours])
            writer.writerow(["RGB colours frequencies", *rgb_frequencies])
            writer.writerow(["RGB colours frequencies (%)", *rgb_frequency_percentages])
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Document Name", file_name])
            writer.writerow(["Data set", data_set])
            writer.writerow(["RGB colours", *rgb_colours])
            writer.writerow(["RGB colours frequencies", *rgb_frequencies])
            writer.writerow(["RGB colours frequencies (%)", *rgb_frequency_percentages])
            writer.writerow('')

if __name__ == "__main__":
    if Path(f"results/3_RGB_value/TEST_0.csv").is_file():
        os.remove(Path(f"results/3_RGB_value/TEST_0.csv"))

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
        rgb_values = get_RGB_value_from_each_run(document)
        rgb_values_list = [bin_rgb_values(rgb) for rgb in rgb_values]
        #rgb_bins = Counter(bin_rgb_values(rgb_value) for rgb_value in rgb_values)
        rgb_bins = {key: Counter(rgb_values_list).get(key, 0) for key in ['black', 'white', 'other']}
        rgb_value_amount = len(rgb_values)

        colours = []
        frequencies = []
        frequency_percentages = []

        # print("Color RGB values for each run element:")
        # print(rgb_values)
        #print(list(Counter(rgb_values).keys()))
        #print(list(Counter(rgb_values).values()))

        print(f"Amount of different run RGB values: {rgb_value_amount}")
        for color, frequency in rgb_bins.items():
            frequency_percent = str(round((frequency / rgb_value_amount) * 100, 2)) #.replace(".", ",")
            colours.append(color)
            frequencies.append(frequency)
            frequency_percentages.append(frequency_percent)
            print(f"RGB Color channel: {color}. Frequency: {frequency} ({frequency_percent}%).")
            
        to_csv(path, data_set[i], colours, frequencies, frequency_percentages)
        i += 1