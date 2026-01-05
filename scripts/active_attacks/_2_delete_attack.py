import re
from pathlib import Path
from docx import Document
from docx.text.run import Run
from docx.oxml.ns import qn
import scripts.active_attacks.unified_active_attack_file as unified_active_attack_file

# Splits text into words, deletes every 10th
def delete_every_nth_word(
        run: Run, 
        n: int, 
        word_index: int
        ) -> int:
    
    text = run.text
    text_pattern = re.compile(r'(\s+)')
    tokens = text_pattern.split(text)
    text_list = []

    for token in tokens:
        if token.strip() == "":
            text_list.append(token)
            continue
        else:
            word_index += 1
            if word_index == (n - 1):
                text_list.append("")
                word_index = 0
            else:
                text_list.append(token)

    run.text = "".join(text_list)
    return word_index

def delete_attack(
        stegoDocPath: str, 
        attackedStegoDocPath: str
        ) -> None:
    
    document = Document(stegoDocPath)
    every_nth_word = 10
    word_index = 0

    print(f"Deleting every {every_nth_word}th word in the {Path(stegoDocPath).name}")
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run_element = run._r
            if run_element.find(qn('w:t')) != None:
                word_index = delete_every_nth_word(run, every_nth_word, word_index)

    # Save
    document.save(attackedStegoDocPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("02_delete_attack", delete_attack, False)