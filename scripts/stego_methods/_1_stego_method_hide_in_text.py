import re
import base64
from copy import deepcopy
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import unified_stego_file

# Count words in paragraphs
def count_words_in_paragraphs(
        document: DocumentObject,
        index: int
    ) -> int:
    word_count = 0
    for paragraph in document.paragraphs[index:]:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        words = re.findall(r'\S+', text, flags=re.UNICODE)
        word_count += len(words)
    # print(f"Total word count: {word_count}")
    return word_count

# Check if max capacity is enough for the message
def is_capacity_enough_for_message(
        document: DocumentObject,
        word_count: int,
        stego_message_toBase64_size_bits:
        int,
        index: int
    ) -> bool:
    # in theory
    # equal_mark = "="
    # equal_mark_count = 0
    # for char in stego_message_toBase64_text:
    #    if char == equal_mark:
    #        equal_mark_count += 1
    # cap = 6 * (word_count - 1 - equal_mark_count)

    # in practice
    text = unified_stego_file.extract_text(document, False, index)
    text_whitespaces = re.findall(r'\x20', text, flags=re.UNICODE)
    cap = len(text_whitespaces)
    is_valid = stego_message_toBase64_size_bits / 6 <= cap
    return is_valid

# Stego-message in Base64
def stego_message_base64(stego_message_bytes: bytes) -> tuple[str, bytes]:
    stego_message_toBase64_bytes = base64.b64encode(stego_message_bytes)
    stego_message_toBase64_text = stego_message_toBase64_bytes.decode('ascii')
    # print(f"Stego-message in Base64: {stego_message_toBase64_text}")
    return stego_message_toBase64_text, stego_message_toBase64_bytes

# Create a new run
def insert_in_run(
        previous_run: Run,
        base_run: Run,
        char: str,
        type: str | None
    ) -> Run:
    current_run_element = previous_run._r
    base_run_element = base_run._r

    # New run element
    new_run_element = OxmlElement('w:r')

    # New run properties
    if base_run_element.rPr != None:
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
            if color_element == None:
                color_element = OxmlElement('w:color')
                run_properties.append(color_element)
            color_element.set(qn('w:val'), 'FFFFFF')

            font_size_element = run_properties.find(qn('w:sz'))
            if font_size_element == None:
                font_size_element = OxmlElement('w:sz')
                run_properties.append(font_size_element)
            font_size_element.set(qn('w:val'), '2')

            vanish_element = run_properties.find(qn('w:vanish'))
            if vanish_element == None:
                vanish_element = OxmlElement('w:vanish')
                run_properties.append(vanish_element)
            vanish_element.set(qn('w:val'), 'true')
        case _:
            text_element.text = char           
    
    current_run_element.addnext(new_run_element)
    new_run = Run(new_run_element, previous_run._parent)
    return new_run

# Splitting the existing runs into before, current and after
def slipt_run_for_embedding(
        run: Run,
        char: str
    ) -> Run | None:
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
def embedding_in_run(
        run: Run,
        stego_message_toBase64_text: str,
        stego_index: int,
        payload: int
    ) -> int:
    current_run = run
    #text = run.text.replace('\xa0', '\x20') # NBSP -> space
    text_whitespaces = re.findall(r'\x20', run.text, flags=re.UNICODE)
    nr_of_unused_whitespace = len(text_whitespaces)

    for _ in range(nr_of_unused_whitespace):
        if stego_index < payload:
            next_run = slipt_run_for_embedding(current_run, stego_message_toBase64_text[stego_index])
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
    stego_message_as_base64 = ''
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
                            stego_message_as_base64 += run.text
    #print(stego_message_as_base64)
    stego_message = base64.b64decode(stego_message_as_base64).decode('utf-8')
    #print(stego_message)
    return stego_message

### Main function ###
def stego_method_hide_in_text() -> None:
    _, stego_message_bytes = unified_stego_file.stego_message()

    unified_stego_file.stego_method(
        'stego_method_1',
        stego_message_base64(stego_message_bytes),
        count_words_in_paragraphs,
        is_capacity_enough_for_message,
        embedding_in_run,
        'string',
        stego_message_extraction
    )

if __name__ == "__main__":
    stego_method_hide_in_text()