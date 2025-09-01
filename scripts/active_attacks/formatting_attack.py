from pathlib import Path
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor

def change_text_style_for_auto_text(document):
    styles = document.styles
    new_style = styles.add_style('stego-style', WD_STYLE_TYPE.CHARACTER)

    new_style_font = new_style.font
    new_style_font.name = 'Britannic Bold'
    new_style_font.size = Pt(16)
    new_style_font.color.rgb = RGBColor(0x23, 0x23, 0x23)

    for paragraph in document.paragraphs:
        paragraph_element = paragraph._p
        paragraph_properties = paragraph_element.pPr
        if paragraph_properties != None:
            pStyle = paragraph_properties.pStyle
        else:
            pStyle = None
        for run in paragraph.runs:
            run_element = run._r
            run_properties = run_element.rPr
            if run_properties != None:
                rStyle = run_properties.rStyle
            else:
                rStyle = None
            if pStyle == None and rStyle == None:
                run.style = new_style

base = "data_set/stego-files"
for directories in Path(base).iterdir():
    for file in directories.iterdir():
        docPath = f"{base}/{directories.name}/{file.name}" #Path("data_set/clean_files/TEST_0.docx")
        document = Document(docPath)
        change_text_style_for_auto_text(document)
        stegoDocPath = Path(f"data_set/attacked_stego-files/4_format_attack/{directories.name}/{file.name}")
        document.save(stegoDocPath)
        print("Saved:", stegoDocPath)