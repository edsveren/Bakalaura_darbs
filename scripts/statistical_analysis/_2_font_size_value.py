from pathlib import Path
from collections import Counter
from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.oxml.styles import CT_Styles
import unified_statistical_analysis_file

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

def document_default_font_size_value(styles_element) -> float | None:
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

def document_default_font_size_pt(document: DocumentObject) -> float:
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
            if run_font_size != None:
                font_sizes.append(run_font_size.pt)
            elif run_font_style_size != None:
                font_sizes.append(run_font_style_size.pt)
            elif paragraph_font_style_size != None:
                font_sizes.append(paragraph_font_style_size.pt)
            else:
                font_sizes.append(default_pt)
    return font_sizes

def font_size_value_analysis(path: Path, data_set: str) -> list[list]:
    document = Document(str(path))
    font_sizes = get_font_size_value_from_each_run(document)
    font_sizes_amount = len(font_sizes)
    counter = Counter(font_sizes)        

    font_sizes_value = list(counter.keys())
    font_sizes_frequencies = list(counter.values())
    frequency_percentages = []
    
    # print(f"Font sizes values for each run element: {font_sizes}")
    # print(f"Font sizes: {font_sizes_value}")
    # print(f"Font size frequencies: {font_sizes_frequencies}")
    for size, frequency in counter.items():
        frequency_percent = str(round((frequency / font_sizes_amount) * 100, 2)) #.replace(".", ",")
        frequency_percentages.append(frequency_percent)
        # print(f"Font size: {size} pt. Frequency: {frequency}")

    data_to_csv = [
        ["Document Name", path.stem],
        ["Data set", data_set],
        ["Font Sizes map", *font_sizes],
        ["Font Sizes (pt)", *font_sizes_value],
        ["Font Sizes Count", *font_sizes_frequencies],
        ["Font Sizes Count (%)", *frequency_percentages]
    ]

    return data_to_csv

def main() -> None:
    unified_statistical_analysis_file.singular_check('2_font_sizes', 'TEST_0', font_size_value_analysis)
        
if __name__ == "__main__":
    main()