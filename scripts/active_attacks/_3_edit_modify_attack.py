import re
from docx import Document
from pathlib import Path
from docx.text.run import Run
from docx.oxml.ns import qn

def change_every_nth_word(run: Run, n: int, word_index: int) -> int:
    text = run.text
    text_pattern = re.compile(r'(\s+)')
    tokens = text_pattern.split(text)
    #print(tokens)
    #print("".join(tokens))
    text_list = []

    for token in tokens:
        if token.strip() == "":
            text_list.append(token)
            continue
        else:
            word_index += 1
            if word_index == (n - 1):
                text_list.append("[[[CHANGED]]]")
                word_index = 0
            else:
                text_list.append(token)
    #print(text_list)
    run.text = "".join(text_list)
    #print(text)
    return word_index

base = "data_set/stego_files"
for directories in Path(base).iterdir():
    for file in directories.iterdir():
        print()
        if file.name.startswith('~$'):
            continue
        docPath = f"{base}/{directories.name}/{file.name}" #Path("data_set/clean_files/TEST_0.docx")
        print("Opened:", docPath)
        print("Attacking:", docPath)
        document = Document(docPath)
        every_nth_word = 10
        word_index = 0
        for paragraph in document.paragraphs:
            for run in paragraph.runs:
                run_element = run._r
                if run_element.find(qn('w:t')) != None:
                    word_index = change_every_nth_word(run, every_nth_word, word_index)

        stegoDocPath = Path(f"data_set/attacked_stego_files/3_edit-modify_attack/{directories.name}/{file.name}")
        document.save(str(stegoDocPath))
        print("Saved:", stegoDocPath)