import re
import random
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.text.run import Run
from docx.oxml.shared import OxmlElement, qn

# Count words in paragraphs
def count_words_in_paragraphs(doc, index) -> int:
    word_count = 0
    for paragraph in doc.paragraphs[index:]:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        words = re.findall(r'\S+', text, flags=re.UNICODE)
        word_count += len(words)
    #print("Kopējais vārdu skaits:", word_count)
    return word_count

# Check if max capacity is enough for the message
def is_capacity_enough_for_message(word_count, stegoMessage_size_bits) -> bool:
    cap = 2 * (word_count - 1)
    is_valid = stegoMessage_size_bits <= cap
    return is_valid

# Extract text from the document
def extract_text(doc) -> str:
    text = []
    for p in doc.paragraphs:
        text.append(p.text.replace('\xa0', '\x20')) # NBSP -> space
    text = "\n".join(text)
    #print("Teksts no dokumenta:\n", text)
    return text

# Stego-message
def stego_message() -> tuple[list[str], bytes]:
    stegoMessagePath = Path("stego-message.txt")
    stegoMessage_bytes = stegoMessagePath.read_bytes()
    stegoMessageText = []
    with open(stegoMessagePath, encoding="utf-8") as input:
        for line in input:
            stegoMessageText.append(line)
    #print("Stego-message data:", stegoMessageText)
    return stegoMessageText, stegoMessage_bytes

# Choose random paragraph
def choose_random_paragraph(doc, stegoMessage_size_bits) -> int | None:
    paragraphs = doc.paragraphs
    if not paragraphs:
        return None
    while True:
        random_paragraph_index = random.randint(0, len(paragraphs) - 1)
        random_paragraph = paragraphs[random_paragraph_index]
        word_count = count_words_in_paragraphs(doc, random_paragraph_index)
        is_valid = is_capacity_enough_for_message(word_count, stegoMessage_size_bits)
        if is_valid:
            print("Random paragraph start:", random_paragraph.text)
            return random_paragraph_index

# Create a new run
def insert_in_run(previous_run, char, type, base_run) -> Run:
    current_run_element = previous_run._r
    base_run_element = base_run._r

    # New run element
    new_run_element = OxmlElement('w:r')

    # New run properties
    if base_run_element.rPr is not None:
        # Copy all existing run properties
        new_run_element.append(deepcopy(base_run_element.rPr))
        run_properties = new_run_element.find(qn('w:rPr'))
    else:
        # Create an empty run properties element
        run_properties = OxmlElement('w:rPr')
        new_run_element.insert(0, run_properties)

    # New text element
    text_element = OxmlElement('w:t')
    new_run_element.append(text_element)

    # Insert character based on type
    match type:
        case 'whitespace':
            text_element.text = char
            #if text_element.text.startswith('\x20') or text_element.text.endswith('\x20'):
            text_element.set(qn('xml:space'), 'preserve')
        case 'stego_char': # len(char) == 1 and char != ('\x20', '\xa0')
            text_element.text = char
            color_element = OxmlElement('w:color')
            font_size_element = OxmlElement('w:sz')
            vanish_element = OxmlElement('w:vanish')
            color_element.set(qn('w:val'), 'FFFFFF')
            font_size_element.set(qn('w:val'), '2')
            vanish_element.set(qn('w:val'), 'true')
            run_properties.append(color_element)
            run_properties.append(font_size_element)
            run_properties.append(vanish_element)
        case _:
            text_element.text = char           
    
    current_run_element.addnext(new_run_element)
    return Run(new_run_element, previous_run._parent)

def slipt_run_for_embedding(run, char) -> Run | None:
    text = run.text
    whitespace = text.find('\x20')
    if whitespace == -1:
        return None

    left_text = text[:whitespace] # text before the first whitespace
    right_text = text[whitespace + 1:] # text after the first whitespace

    # left text
    run.text = left_text

    # left whitespace
    left_whitespace = insert_in_run(run, '\x20', 'whitespace', run)

    # stego character
    stego_char = insert_in_run(left_whitespace, char, 'stego_char', run)

    # right whitespace
    right_whitespace = insert_in_run(stego_char, '\x20', 'whitespace', run)

    # right text
    remaining_run = insert_in_run(right_whitespace, right_text, None, run)
    return remaining_run

def embedding_in_run(run, stegoMessage_toBase64_text, stego_index, payload) -> int:
    current_run = run
    #text = run.text.replace('\xa0', '\x20') # NBSP -> space
    text_whitespaces = re.findall(r'\x20', run.text, flags=re.UNICODE)
    nr_of_unused_whitespace = len(text_whitespaces)

    for _ in range(nr_of_unused_whitespace):
        if stego_index < payload:
            next_run = slipt_run_for_embedding(current_run, stegoMessage_toBase64_text[stego_index])
            if next_run != None:
                current_run = next_run
                stego_index += 1
            else:
                break
        else:
            break
    return stego_index
            
# DOCX file
docPath = Path("TEST_CASES\TEST_0\TEST_0.docx")
doc = Document(docPath)
text = extract_text(doc)
word_count = count_words_in_paragraphs(doc, 0)

stego_message_text, stegoMessage_bytes = stego_message()
stegoMessage_size_bytes = len(stegoMessage_bytes)
stegoMessage_size_bits = 8 * stegoMessage_size_bytes
#print("Regular bytes:", stegoMessage_size_bytes)
#print("Regular bites:", stegoMessage_size_bits)

# Main

embedded = False
while not embedded:
    # Check if the paragraph has enough runs to embed the message
    is_valid = is_capacity_enough_for_message(word_count, stegoMessage_size_bits)
    print("The cover object is valid:", is_valid)
    if not is_valid:
        print("Not enough capacity in the document to embed the message.")
        break

    random_paragraph_index = choose_random_paragraph(doc, stegoMessage_size_bits)
    if random_paragraph_index is None:
        print("No paragraphs available for embedding.")
        break

    # Embed stego-message in DOCX
    print("Embedding stego-message...")
    payload = len(stego_message_text)
    stego_index = 0
    #while payload < stego_index:
    for paragraph in doc.paragraphs [random_paragraph_index:]:
            if stego_index < payload:
                original_run_amount = list(paragraph.runs)
                for run in original_run_amount:
                    if stego_index < payload:
                        stego_index = embedding_in_run(run, stegoMessage_toBase64_text, stego_index, payload)
                    else:
                        break
            else:
                break

    embedded = True
    print("Embedding successful!")


if embedded:
    stegoDocPath = docPath.with_name(docPath.stem + "_STEGO.docx")
    doc.save(stegoDocPath)
    print("Saved:", stegoDocPath)
else:
    print("Embedding not possible.")