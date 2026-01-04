from pathlib import Path
from collections import Counter
from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.oxml.styles import CT_Styles
import scripts.statistical_analysis.unified_statistical_analysis_file as unified_statistical_analysis_file

# Find the default font size in a paragraph element
def paragraph_default_font_size_value(styles_element: CT_Styles) -> float | None:
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

# Find the default font size in the entire document
def document_default_font_size_value(styles_element: CT_Styles) -> float | None:
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

# Find the default font size value
def document_default_font_size_pt(document: DocumentObject) -> float:
    styles_element = document.styles.element

    font_size_value_pt = 11.0
    # paragraph_font_size_value_pt = styles_element.xpath(".//w:style[@w:type='paragraph' and @w:default='1']//w:rPr/w:sz/@w:val")
    # default_font_size_value_pt = styles_element.xpath(".//w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")

    paragraph_font_size_value_pt = paragraph_default_font_size_value(styles_element)
    default_font_size_value_pt = document_default_font_size_value(styles_element)
    
    # If there is a universal default font size
    if paragraph_font_size_value_pt:
        # font_size_value_pt = int(paragraph_font_size_value_pt[0]) / 2.0
        font_size_value_pt = paragraph_font_size_value_pt
        return font_size_value_pt
    # If there is a default font size in a paragraph element
    elif default_font_size_value_pt:
        # font_size_value_pt = int(default_font_size_value_pt[0]) / 2.0
        font_size_value_pt = default_font_size_value_pt
        return font_size_value_pt
    # Otherwise, the default font size is always set as 11
    else:
        return font_size_value_pt

# Get the current font size value in each run
def get_font_size_value_from_each_run(document: DocumentObject) -> list:
    default_pt = document_default_font_size_pt(document)
    font_sizes = []
    for paragraph in document.paragraphs:
        if paragraph.style != None:
            paragraph_font_style_size = paragraph.style.font.size
        else:
            paragraph_font_style_size = None
            
        for run in paragraph.runs:
            run_font_size = run.font.size
            run_font_style_size = run.style.font.size

            # If the current font size is set at the run level
            if run_font_size != None:
                font_sizes.append(run_font_size.pt)
            # If the current font size is set at the run style level
            elif run_font_style_size != None:
                font_sizes.append(run_font_style_size.pt)
            # If the current font size is set at the paragraph style level
            elif paragraph_font_style_size != None:
                font_sizes.append(paragraph_font_style_size.pt)
            # Otherwise the current font size is default
            else:
                font_sizes.append(default_pt)
    return font_sizes

# Bin the entire font size list into 5 categories
def bin_font_size_values(font_size: float) -> str:
    if font_size > 20:
        return '.+20.'
    elif font_size > 11:
        return '.20-11.'
    elif font_size == 11:
        return '.11.'
    elif font_size > 1:
        return '.10-1.'
    else:
        return '.1.'

### Main function ###
def font_size_value_analysis(path: Path, data_set: str, chosen_file: bool) -> tuple[list[list], int]:
    document = Document(str(path))
    font_size_values = get_font_size_value_from_each_run(document)
    font_sizes_amount = len(font_size_values)

    ## For chosen document
    counter = Counter(font_size_values)
    counter_sorted = dict(sorted(counter.items(), reverse=True))        

    font_sizes_value = list(counter_sorted.keys())
    font_sizes_frequencies = list(counter_sorted.values())
    frequency_percentages = []
    
    # print(f"Font sizes values for each run element: {font_sizes}")
    # print(f"Font sizes: {font_sizes_value}")
    # print(f"Font size frequencies: {font_sizes_frequencies}")
    for size, frequency in counter_sorted.items():
        frequency_percent = str(round((frequency / font_sizes_amount) * 100, 2))
        frequency_percentages.append(frequency_percent)
        # print(f"Font size: {size} pt. Frequency: {frequency}")

    ## For all documents
    font_sizes_list = [bin_font_size_values(font_size) for font_size in font_size_values]
    font_sizes_bins = {key: Counter(font_sizes_list).get(key, 0) for key in ['.+20.', '.20-11.', '.11.', '.10-1.', '.1.']}

    font_sizes_value_multi = []
    font_sizes_frequencies_multi = []
    frequency_percentages_multi = []

    for size, frequency in font_sizes_bins.items():
        frequency_percent = str(round((frequency / font_sizes_amount) * 100, 2))
        font_sizes_value_multi.append(size)
        font_sizes_frequencies_multi.append(frequency)
        frequency_percentages_multi.append(frequency_percent)

    # Data export
    if not chosen_file:
        data_to_csv = [
            ["Data set", data_set],
            ["Document Name", "Run Font Sizes (pt) frequencies", '', '', '', '', "Run Font Sizes (pt) frequencies (%)", '', '', '', ''],
            ['', *font_sizes_value_multi, *font_sizes_value_multi],
            [path.stem, *font_sizes_frequencies_multi, *frequency_percentages_multi]
        ]
    else:
        data_to_csv = [
            ["Document Name", path.stem],
            ["Data set", data_set],
            ["Font Sizes map", *font_size_values],
            ["Font Sizes (pt)", *font_sizes_value],
            ["Font Sizes Count", *font_sizes_frequencies],
            ["Font Sizes Count (%)", *frequency_percentages]
        ]

    return data_to_csv, 3

def main() -> None:
    unified_statistical_analysis_file.statistical_analysis('2_font_sizes', 'TEST_0', font_size_value_analysis)
    unified_statistical_analysis_file.statistical_analysis('2_font_sizes', None, font_size_value_analysis)
        
if __name__ == "__main__":
    main()