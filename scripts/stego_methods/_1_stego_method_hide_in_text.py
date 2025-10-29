import re
import base64
import random
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject

# Count words in paragraphs
def count_words_in_paragraphs(document: DocumentObject, index: int) -> int:
    word_count = 0
    for paragraph in document.paragraphs[index:]:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        words = re.findall(r'\S+', text, flags=re.UNICODE)
        word_count += len(words)
    #print("Kopējais vārdu skaits:", word_count)
    return word_count

# Check if max capacity is enough for the message
def is_capacity_enough_for_message(document: DocumentObject, stegoMessage_toBase64_size_bits: int, index: int=0) -> bool:
    # in theory
    # equal_mark = "="
    # equal_mark_count = 0
    # for char in stegoMessage_toBase64_text:
    #    if char == equal_mark:
    #        equal_mark_count += 1
    # cap = 6 * (word_count - 1 - equal_mark_count)

    # in practice
    text = extract_text(document, False, index)
    text_whitespaces = re.findall(r'\x20', text, flags=re.UNICODE)
    cap = len(text_whitespaces)
    is_valid = stegoMessage_toBase64_size_bits / 6 <= cap
    return is_valid

# Extract text from the document
def extract_text(document: DocumentObject, NBSP: bool, index: int=0) -> str:
    text = []
    for paragraph in document.paragraphs[index:]:
        if NBSP == True:
            text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
        else:
            text.append(paragraph.text)
    text = "\n".join(text)
    text = ''.join(text)
    #print("Teksts no dokumenta:\n", text)
    return text

# Stego-message
def stego_message() -> tuple[str, bytes]:
    stegoMessageText = Path("stego_messages\stego_message.txt").read_text(encoding="utf-8")
    stegoMessage_bytes = stegoMessageText.encode("utf-8")
    return stegoMessageText, stegoMessage_bytes

# Stego-message in Base64
def stego_message_base64(stegoMessage_bytes: bytes) -> tuple[str, bytes]:
    stegoMessage_toBase64_bytes = base64.b64encode(stegoMessage_bytes)
    stegoMessage_toBase64_text = stegoMessage_toBase64_bytes.decode('ascii')
    #print("Stego-message Base64:", stegoMessage_toBase64_text)
    return stegoMessage_toBase64_text, stegoMessage_toBase64_bytes

# Choose random paragraph
def choose_random_paragraph(document: DocumentObject, stegoMessage_toBase64_size_bits: int) -> int | None:
    paragraphs = document.paragraphs
    if not paragraphs:
        return None
    while True:
        random_paragraph_index = random.randint(0, len(paragraphs) - 1)
        random_paragraph = paragraphs[random_paragraph_index]
        word_count = count_words_in_paragraphs(document, random_paragraph_index)
        is_valid = is_capacity_enough_for_message(document, stegoMessage_toBase64_size_bits, random_paragraph_index)
        if is_valid:
            #print("Random paragraph start:", random_paragraph.text)
            return random_paragraph_index

# Create a new run
def insert_in_run(previous_run: Run, base_run: Run, char: str, type: str | None) -> Run:
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

            color_element = run_properties.find(qn('w:color'))
            if color_element is None:
                color_element = OxmlElement('w:color')
                run_properties.append(color_element)
            color_element.set(qn('w:val'), 'FFFFFF')

            font_size_element = run_properties.find(qn('w:sz'))
            if font_size_element is None:
                font_size_element = OxmlElement('w:sz')
                run_properties.append(font_size_element)
            font_size_element.set(qn('w:val'), '2')

            vanish_element = run_properties.find(qn('w:vanish'))
            if vanish_element is None:
                vanish_element = OxmlElement('w:vanish')
                run_properties.append(vanish_element)
            vanish_element.set(qn('w:val'), 'true')
        case _:
            text_element.text = char           
    
    current_run_element.addnext(new_run_element)
    new_run = Run(new_run_element, previous_run._parent)
    return new_run

# Splitting the existing runs into before, current and after
def slipt_run_for_embedding(run: Run, char: str) -> Run | None:
    text = run.text
    whitespace = text.find('\x20')
    if whitespace == -1:
        return None

    left_text = text[:whitespace] # text before the first whitespace
    right_text = text[whitespace + 1:] # text after the first whitespace

    # left text
    run.text = left_text

    # left whitespace
    left_whitespace = insert_in_run(run, run, '\x20', 'whitespace')

    # stego character
    stego_char_run = insert_in_run(left_whitespace, run, char, 'stego_char')

    # right whitespace
    right_whitespace = insert_in_run(stego_char_run, run, '\x20', 'whitespace')

    # right text
    remaining_run = insert_in_run(right_whitespace, run, right_text, None)
    return remaining_run

# Embedding algorithm
def embedding_in_run(run: Run, stegoMessage_toBase64_text: str, stego_index: int, payload: int) -> int:
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

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    stegoMessage_as_base64 = ''
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run_properties = run._r.rPr
            if run_properties != None:
                if len(run.text) == 1:
                    #base_64_char = re.search(r'[A-Za-z0-9+/]', run.text, flags=re.UNICODE)
                    color_element = run_properties.find(qn('w:color'))
                    font_size_element = run_properties.find(qn('w:sz'))
                    vanish_element = run_properties.find(qn('w:vanish'))

                    if None not in (color_element, font_size_element, vanish_element):
                        color_element_value = color_element.get(qn('w:val'))
                        font_size_value = font_size_element.get(qn('w:val'))
                        #vanish_element_value = vanish_element.get(qn('w:val'))
                        if color_element_value.upper() == 'FFFFFF' and font_size_value == '2': #and base_64_char != None:
                            stegoMessage_as_base64 += run.text
    #print(stegoMessage_as_base64)
    stegoMessage = base64.b64decode(stegoMessage_as_base64).decode('utf-8')
    #print(stegoMessage)
    return stegoMessage

### Main ###      
# DOCX file
base = "data_set/clean_files"
for file in Path(base).iterdir():
    docPath = f"{base}/{file.name}" #Path("data_set/clean_files/TEST_0.docx")
    print(f"DOCX file: {docPath}")
    print("Beginning the embedding process...")
    document = Document(docPath)
    text = extract_text(document, True)
    word_count = count_words_in_paragraphs(document, 0)

    stego_message_text, stegoMessage_bytes = stego_message()
    stegoMessage_size_bytes = len(stegoMessage_bytes)
    stegoMessage_size_bits = 8 * stegoMessage_size_bytes
    #print("Regular bytes:", stegoMessage_size_bytes)
    #print("Regular bites:", stegoMessage_size_bits)

    stegoMessage_toBase64_text, stegoMessage_toBase64_bytes = stego_message_base64(stegoMessage_bytes)
    stegoMessage_toBase64_size_bytes = len(stegoMessage_toBase64_bytes)
    stegoMessage_toBase64_size_bits = 8 * stegoMessage_toBase64_size_bytes
    #print("Stego-message Base64 bytes:", stegoMessage_toBase64_size_bytes)
    #print("Stego-message Base64 bits:", stegoMessage_toBase64_size_bits)

    embedded = False
    while not embedded:
        # Check if the paragraph has enough runs to embed the message
        print("Checking if the cover object is valid for embedding...")
        is_valid = is_capacity_enough_for_message(document, stegoMessage_toBase64_size_bits)
        print("The cover object is valid:", is_valid)
        if not is_valid:
            print("Not enough capacity in the document to embed the message.")
            break

        # Embed stego-message in DOCX
        print("Embedding stego-message...")
        random_paragraph_index = choose_random_paragraph(document, stegoMessage_toBase64_size_bits)
        if random_paragraph_index is None:
            print("No paragraphs available for embedding.")
            break

        payload = stegoMessage_toBase64_size_bytes
        stego_index = 0
        #while payload < stego_index:
        for paragraph in document.paragraphs [random_paragraph_index:]:
            if stego_index < payload:
                original_run_amount = list(paragraph.runs)
                for run in original_run_amount:
                    run_element = run._r
                    # Only process runs that contain text
                    if run_element.find(qn('w:t')) != None:
                        if stego_index < payload:
                            next_stego_index = embedding_in_run(run, stegoMessage_toBase64_text, stego_index, payload)
                            stego_index = next_stego_index
                        else:
                            break
            else:
                break
        
        #print("Extracting stego-message...")
        if stego_message_text != stego_message_extraction(document):
            print("Extracted message is not equal to stego-message!")
            break
        #print("Extraction successful!")
        print("Embedding successful!")
        embedded = True

    if embedded:
        stegoDocPath = str(Path(f"data_set/stego_files/stego_method_1/{file.name}"))
        document.save(stegoDocPath)
        print("Saved:", stegoDocPath)
    else:
        print("Embedding not possible.")
    print()
