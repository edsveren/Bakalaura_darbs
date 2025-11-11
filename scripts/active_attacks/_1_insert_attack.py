import re
from pathlib import Path
from docx import Document
from docx.text.run import Run
from docx.oxml.ns import qn
import unified_active_attack_file

def insert_word_at_every_nth_word(
        run: Run, 
        n: int, 
        word_index: int
        ) -> int:
    
    # Get the entire run text
    text = run.text
    # A ReGex that finds blocks of whitespace and preserves them as individual tokens
    text_pattern = re.compile(r'(\s+)')
    # Split text into words and whitespaces
    tokens = text_pattern.split(text)
    # A list that will be used to reform the text
    text_list = []

    # Loop through each word and whitespace
    for token in tokens:
        # If the current text token is whitespace, just add it to the list
        if token.strip() == "":
            text_list.append(token)
            continue
        else:
            # Increase index after each found word token
            word_index += 1
            # When the index is equal to the index for insertion
            if word_index == (n - 1):
                # Insert a word marker
                text_list.append("[[[INSERTED]]]")
                # Add the word
                text_list.append(" ")
                text_list.append(token)
                # Reset the index
                word_index = 0
            # Otherwise, just add word the list
            else:
                text_list.append(token)

    # Replace the current run text with the created list
    run.text = "".join(text_list)

    # Return current index
    return word_index

def insert_attack(
        stegoDocPath: str, 
        attackedStegoDocPath: str
        ) -> None:
    
    # Open the stego-file
    document = Document(stegoDocPath)
    # Index of when to insert a word
    every_nth_word = 10
    # Index to follow when to insert a word
    word_index = 0

    print(f"Inserting a word after every {every_nth_word}th word in the {Path(stegoDocPath).name}")
    # Go through each run in each paragraph
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            # Get the run element and attack it if it has at least one text element 
            run_element = run._r
            if run_element.find(qn('w:t')) != None:
                word_index = insert_word_at_every_nth_word(run, every_nth_word, word_index)

    # Save
    document.save(attackedStegoDocPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("01_insert_attack", insert_attack, False)