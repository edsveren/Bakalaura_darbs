from docx import Document
from pathlib import Path

def count_text_elements(document) -> int:
    root = document.part.element
    docTextElement = root.xpath("//w:document/w:body/w:p/w:r/w:t")
    return len(docTextElement)

def document_default_font_size_pt(document) -> int:
    styles = document.styles.element

    font_size_value_pt = 11.0
    paragraph_font_size_value_pt = styles.xpath(".//w:style[@w:type='paragraph' and @w:default='1']//w:rPr/w:sz/@w:val")
    default_font_size_value_pt = styles.xpath(".//w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")

    if paragraph_font_size_value_pt:
        font_size_value_pt = int(paragraph_font_size_value_pt[0]) / 2.0
        return font_size_value_pt
    elif default_font_size_value_pt:
        font_size_value_pt = int(default_font_size_value_pt[0]) / 2.0
        return font_size_value_pt
    else:
        return font_size_value_pt

def get_font_size_value_from_each_run(document) -> list:
    default_pt = document_default_font_size_pt(document)
    font_sizes = []
    for paragraph in document.paragraphs:
        paragraph_font_style_size = paragraph.style.font.size
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
    with open("results/font_sizes.txt", "w", encoding="utf-8") as file:
        for size in font_sizes:
            file.write(str(size) + "\n")
    return font_sizes

docPath_0 = Path("data_set/clean_files/TEST_0.docx")
docPath_1 = Path("data_set/stego-files/stego-method_1/TEST_0.docx")
docPath_2 = Path("data_set/stego-files/stego-method_4/TEST_0.docx")
docPath_3 = Path("data_set/stego-files/stego-method_5/TEST_0.docx")
docPath_4= Path("data_set/stego-files/stego-method_6/TEST_0.docx")

paths = [docPath_0, docPath_1, docPath_2, docPath_3, docPath_4]

for path in paths:
    print("")
    print(f"DOCUMENT: {path}")
    document = Document(path)
    print(f"Text element count: {count_text_elements(document)}")
    print("Font sizes values for each run element:")
    print(get_font_size_value_from_each_run(document))