import re
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import unified_stego_file

zero = '\u2009' # Thin space
one = '\u200A'  # Hair space
two = '\u200B'  # Zero width space

unispace_dictionary = {
    (zero, zero, zero): 'A',
    (zero, zero, one): 'B',
    (zero, zero, two): 'C',
    (zero, one, zero): 'D',
    (zero, one, one): 'E',
    (zero, one, two): 'F',
    (zero, two, zero): 'G',
    (zero, two, one): 'H',
    (zero, two, two): 'I',
    (one, zero, zero): 'J',
    (one, zero, one): 'K',
    (one, zero, two): 'L',
    (one, one, zero): 'M',
    (one, one, one): 'N',
    (one, one, two): 'O',
    (one, two, zero): 'P',
    (one, two, one): 'Q',
    (one, two, two): 'R',
    (two, zero, zero): 'S',
    (two, zero, one): 'T',
    (two, zero, two): 'U',
    (two, one, zero): 'V',
    (two, one, one): 'W',
    (two, one, two): 'X',
    (two, two, zero): 'Y',
    (two, two, one): 'Z',
    (two, two, two): '\x20', # space
}

reverse_unispace_dictionary = {value: key for key, value in unispace_dictionary.items()}

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
        stego_message_size_bits: int,
        index: int
    ) -> bool:
    cap = 8 * (word_count - 1)
    is_valid = stego_message_size_bits <= cap
    return is_valid

def stego_message_standarization_to_unispace_method(stego_messageText: str) -> str:
    stego_message_in_whitespace_unicode = ''
    for line in stego_messageText:
        for char in line:
            char = char.upper()
            key = reverse_unispace_dictionary.get(char)
            if key != None:
                stego_message_in_whitespace_unicode += char
    return stego_message_in_whitespace_unicode

# Embedding algorithm
def embedding_in_run(
        run: Run,
        stego_message_text: str,
        stego_index: int,
        payload: int
    ) -> int:
    run_element = run._r
    text_whitespaces = re.findall(r'\x20', run.text, flags=re.UNICODE)
    nr_of_unused_whitespace = len(text_whitespaces)

    text = run.text
    embedded_text = ''
    remaining_text = text

    for _ in range(nr_of_unused_whitespace):
        if stego_index >= payload:
            break

        whitespace = re.search(r'\x20', remaining_text, flags=re.UNICODE)
        if whitespace == None:
            embedded_text += remaining_text
            remaining_text = ''
            break
        else:
            cover_whitespace_index = whitespace.start()
            embedded_text += remaining_text[:cover_whitespace_index]
            remaining_text = remaining_text[cover_whitespace_index + 1:]

            # Get the unispace combination for the stego character, default to empty tuple if not found (primarily for linter)
            unispace_combination = reverse_unispace_dictionary.get(stego_message_text[stego_index], ())
            unispace_combination_string = ''.join(unispace_combination)
            embedded_text += unispace_combination_string
            stego_index += 1

    if remaining_text != '':
        embedded_text += remaining_text

    if embedded_text != text:
        run.text = embedded_text
        run_properties = run_element.rPr
        if run_properties == None:
            run_properties = OxmlElement("w:rPr")
            run_element.insert(0, run_properties)
        if run_properties.find(qn("w:noProof")) == None:
            run_properties.append(OxmlElement("w:noProof"))

        for text_element in run_element.iter(qn('w:t')):
            text_element.set(qn('xml:space'), 'preserve')

    return stego_index

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    text = unified_stego_file.extract_text(document, True, 0)
    unispace_combination_as_string_dictionary = {''.join(key): value for key, value in unispace_dictionary.items()}
    stego_message = ''
    first_zero_width_space = re.search(two, text, flags=re.UNICODE)
    if first_zero_width_space == None:
        return ''
    cover_text_with_stego = text[first_zero_width_space.start():]

    first_whitespace = re.search(r'\x20', cover_text_with_stego, flags=re.UNICODE)
    if first_whitespace != None:
        cover_text_with_stego = cover_text_with_stego[:first_whitespace.start()]

    unispace_combination_length = 0
    unispace_combination = ''
    for char in cover_text_with_stego:
        if char in (zero, one, two):
            unispace_combination += char
            unispace_combination_length += 1
            if unispace_combination_length == 3:
                stego_char = unispace_combination_as_string_dictionary.get(unispace_combination)
                if stego_char != None:
                    stego_message += stego_char
                    unispace_combination_length = 0
                    unispace_combination = ''
    #print(stego_message)
    stego_message = stego_message[1:-1]
    return stego_message

def stego_method_unispace() -> None:
    stego_message_text, stego_message_bytes = unified_stego_file.stego_message()
    stego_message_in_whitespace_unicode = stego_message_standarization_to_unispace_method(stego_message_text)
    stego_message_in_whitespace_unicode_embedding_ready_format = '\x20' + stego_message_in_whitespace_unicode + '\x20'

    unified_stego_file.stego_method(
        'stego_method_5',
        (stego_message_in_whitespace_unicode_embedding_ready_format, stego_message_bytes),
        count_words_in_paragraphs,
        is_capacity_enough_for_message,
        embedding_in_run,
        'string',
        stego_message_extraction
    )

if __name__ == "__main__":
    stego_method_unispace()