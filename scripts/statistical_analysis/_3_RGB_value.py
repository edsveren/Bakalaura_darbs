from pathlib import Path
from collections import Counter
from docx import Document
from docx.document import Document as DocumentObject
import scripts.statistical_analysis.unified_statistical_analysis_file as unified_statistical_analysis_file

# Get the font RGB colour values from each run element
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

# Bin the entire RGB colour list into 3 categories
def bin_rgb_values(rgb_value: str) -> str:
    match rgb_value:
        case "#000000":   
            return "black"
        case "#FFFFFF":
            return "white"
        case _:
            return "other"

### Main function ###
def RGB_value_analysis(path: Path, data_set: str, chosen_file: bool) -> tuple[list[list], int]:
    document = Document(str(path))
    rgb_values = get_RGB_value_from_each_run(document)
    rgb_values_list = [bin_rgb_values(rgb) for rgb in rgb_values]
    rgb_bins = {key: Counter(rgb_values_list).get(key, 0) for key in ['black', 'white', 'other']}
    rgb_value_amount = len(rgb_values)

    colours = []
    frequencies = []
    frequency_percentages = []

    # print("Color RGB values for each run element:")
    # print(rgb_values)
    # print(list(Counter(rgb_values).keys()))
    # print(list(Counter(rgb_values).values()))

    # print(f"Amount of different run RGB values: {rgb_value_amount}")
    for color, frequency in rgb_bins.items():
        frequency_percent = str(round((frequency / rgb_value_amount) * 100, 2))
        colours.append(color)
        frequencies.append(frequency)
        frequency_percentages.append(frequency_percent)
        # print(f"RGB Color channel: {color}. Frequency: {frequency} ({frequency_percent}%).")

    # Data export
    if not chosen_file:
        data_to_csv = [
            ["Data set", data_set],
            ["Document Name", "Run RGB colours frequencies", '', '', "Run RGB colours frequencies (%)", '', ''],
            ['', *colours, *colours],
            [path.stem, *frequencies, *frequency_percentages]
        ]
    else:
        data_to_csv = [
            ["Document Name", path.stem],
            ["Data set", data_set],
            ["RGB colours", *colours],
            ["RGB colours frequencies", *frequencies],
            ["RGB colours frequencies (%)", *frequency_percentages]
        ]
    
    return data_to_csv, 3

def main() -> None:
    unified_statistical_analysis_file.statistical_analysis('3_RGB_value', 'TEST_0', RGB_value_analysis)
    unified_statistical_analysis_file.statistical_analysis('3_RGB_value', None, RGB_value_analysis)

if __name__ == "__main__":
    main()