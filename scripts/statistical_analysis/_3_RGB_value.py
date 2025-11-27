from pathlib import Path
from collections import Counter
from docx import Document
from docx.document import Document as DocumentObject
import unified_statistical_analysis_file

def get_RGB_value_from_each_run(document: DocumentObject) -> list:
    rgb_values = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color
            if font_color.rgb != None:
                rgb_values.append(f"#{font_color.rgb}")
            else:
                rgb_values.append("#000000")
    return rgb_values

def bin_rgb_values(rgb_value: str) -> str:
    match rgb_value:
        case "#000000":   
            return "black"
        case "#FFFFFF":
            return "white"
        case _:
            return "other"
        
def RGB_value_analysis(path: Path, data_set: str) -> list[list]:
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

    # print(f"Amount of different run RGB values: {rgb_value_amount}")
    for color, frequency in rgb_bins.items():
        frequency_percent = str(round((frequency / rgb_value_amount) * 100, 2)) #.replace(".", ",")
        colours.append(color)
        frequencies.append(frequency)
        frequency_percentages.append(frequency_percent)
        # print(f"RGB Color channel: {color}. Frequency: {frequency} ({frequency_percent}%).")

    data_to_csv = [
        ["Document Name", path.stem],
        ["Data set", data_set],
        ["RGB colours", *colours],
        ["RGB colours frequencies", *frequencies],
        ["RGB colours frequencies (%)", *frequency_percentages]
    ]
    
    return data_to_csv

def main() -> None:
    unified_statistical_analysis_file.singular_check('3_RGB_value', 'TEST_0', RGB_value_analysis)

if __name__ == "__main__":
    main()