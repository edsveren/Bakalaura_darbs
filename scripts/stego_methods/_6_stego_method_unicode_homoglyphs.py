import re
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import unified_stego_file

unicode_dictionary = {
    'a': '\u0430',
    'b': '\u042C',
    'c': '\u03F2',
    'd': '\u0501',
    'e': '\u0435',
    'f': '\uAB35',
    'g': '\u0261',
    'h': '\u04BB',
    'i': '\u0456',
    'j': '\u03F3',
    'k': '\u043A',
    'l': '\u04CF',
    'm': '\u043C',
    'n': '\u0578',
    'o': '\u03BF',
    'p': '\u0440',
    'q': '\u051B',
    'r': '\u1D26',
    's': '\u0455',
    't': '\u03C4',
    'u': '\u057D',
    'v': '\u1D20',
    'w': '\u051D',
    'x': '\u0445',
    'y': '\u0443',
    'z': '\u1D22'
}

reverse_unicode_dictionary = {value: key for key, value in unicode_dictionary.items()}

# Count chars in paragraphs
def count_chars_in_paragraphs(
        document: DocumentObject,
        index: int
    ) -> int:
    char_count = 0
    for paragraph in document.paragraphs[index:]:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        chars = re.findall(r'[a-z]', text, flags=re.UNICODE)
        char_count += len(chars)
    # print(f"Total char count: {char_count}")
    return char_count

# Check if max capacity is enough for the message
def is_capacity_enough_for_message(
        document: DocumentObject,
        char_count: int,
        stego_message_size_bits: int,
        index: int
    ) -> bool:
    cap = char_count
    is_valid = stego_message_size_bits <= cap
    return is_valid

# Transform stego-message to bit string
def stego_message_to_bit_string(stego_message_bytes: bytes) -> str:
    stego_byte_to_binary_string = ''
    for byte in stego_message_bytes:
        stego_byte_to_binary_string += f"{byte:08b}"
    #print(len(stego_byte_to_binary_string))
    return stego_byte_to_binary_string

# Embedding algorithm
def embedding_in_run(
        run: Run,
        stego_message_text: str,
        stego_index: int,
        payload: int
    ) -> int:
    run_element = run._r
    run_properties = run_element.rPr
    if run_properties == None:
        run_properties = OxmlElement("w:rPr")
        run_element.insert(0, run_properties)
    if run_properties.find(qn("w:noProof")) == None:
        run_properties.append(OxmlElement("w:noProof"))

    all_lowercase_letters = re.findall(r'[a-z]', run.text, flags=re.UNICODE)
    nr_of_unused_lowercase_letters = len(all_lowercase_letters)

    text = run.text
    embedded_text = ''
    remaining_text = text

    for _ in range(nr_of_unused_lowercase_letters):
        if stego_index >= payload:
            break

        lowercase_letter = re.search(r'[a-z]', remaining_text, flags=re.UNICODE)
        if lowercase_letter == None:
            embedded_text += remaining_text
            remaining_text = ''
            break
        else:
            cover_symbol_index = lowercase_letter.start()
            single_cover_char = remaining_text[cover_symbol_index]
            embedded_text += remaining_text[:cover_symbol_index]
            remaining_text = remaining_text[cover_symbol_index + 1:]
            if stego_message_text[stego_index] == '1':
                # Get the homoglyph for the stego character, default to the original character if not found (primarily for linter)
                homoglyph = unicode_dictionary.get(single_cover_char, single_cover_char)
                embedded_text += homoglyph
            else:
                embedded_text += single_cover_char
            stego_index += 1

    if remaining_text != '':
        embedded_text += remaining_text

    if embedded_text != text:
        run.text = embedded_text

    return stego_index

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    text = unified_stego_file.extract_text(document, True, 0)
    stego_message_bytes = b''
    first_stego_char_found = False
    stego_byte_string = ''
    end_byte_string = '00000000'
    for char in text:
        stego_char = reverse_unicode_dictionary.get(char)
        if not first_stego_char_found:
            if stego_char != None:
                first_stego_char_found = True
            continue
        else:
            if stego_char != None:
                stego_byte_string += '1'
            elif 'a' <= char <= 'z':
                stego_byte_string += '0'
            else:
                continue
            if len(stego_byte_string) == 8:
                if stego_byte_string == end_byte_string:
                    break
                else:
                    stego_byte = int(stego_byte_string, 2).to_bytes(1, 'big')
                    stego_message_bytes += stego_byte
                    stego_byte_string = ''
    stego_message = stego_message_bytes.decode('utf-8')
    #print(stego_message)
    return stego_message

#### Main function ###
def stego_method_unicode_homoglyphs() -> None:
    _, stego_message_bytes = unified_stego_file.stego_message()
    stego_message_bytes_to_binary_string = '1' + stego_message_to_bit_string(stego_message_bytes)

    unified_stego_file.stego_method(
        'stego_method_6',
        (stego_message_bytes_to_binary_string, stego_message_bytes),
        count_chars_in_paragraphs,
        is_capacity_enough_for_message,
        embedding_in_run,
        'string',
        stego_message_extraction
    )

if __name__ == "__main__":
    stego_method_unicode_homoglyphs()