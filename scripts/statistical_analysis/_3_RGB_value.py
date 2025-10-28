from docx import Document
from pathlib import Path

def count_text_elements(document) -> int:
    root = document.part.element
    docTextElement = root.xpath("//w:document/w:body/w:p/w:r/w:t")
    return len(docTextElement)

def get_RGB_value_from_each_run(document) -> list:
    rgb_values = []
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            font_color = run.font.color
            if font_color.rgb != None:
                rgb_values.append(f"#{font_color.rgb}")
            else:
                rgb_values.append("#000000")
    with open("results/rgb_values.txt", "w", encoding="utf-8") as file:
        for rgb in rgb_values:
            file.write(str(rgb) + "\n")
    return rgb_values

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
    print("Color RGB values for each run element:")
    print(get_RGB_value_from_each_run(document))
