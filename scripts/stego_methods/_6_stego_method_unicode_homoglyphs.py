import re
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import scripts.stego_methods.unified_stego_file as unified_stego_file

# Unicode homoglyph dictionary
unicode_dictionary = {
    'a': '\u0430',  # а
    'b': '\u042C',  # Ь
    'c': '\u03F2',  # ϲ
    'd': '\u0501',  # ԁ
    'e': '\u0435',  # е
    'f': '\uAB35',  # ꬵ
    'g': '\u0261',  # ɡ
    'h': '\u04BB',  # һ
    'i': '\u0456',  # і
    'j': '\u03F3',  # ϳ
    'k': '\u043A',  # к
    'l': '\u04CF',  # ӏ
    'm': '\u043C',  # м
    'n': '\u0578',  # ո
    'o': '\u03BF',  # ο
    'p': '\u0440',  # р
    'q': '\u051B',  # ԛ
    'r': '\u1D26',  # ᴦ
    's': '\u0455',  # ѕ
    't': '\u03C4',  # τ
    'u': '\u057D',  # ս
    'v': '\u1D20',  # ᴠ
    'w': '\u051D',  # ԝ
    'x': '\u0445',  # х
    'y': '\u0443',  # у
    'z': '\u1D22'   # ᴢ
}

# Reverse unicode homoglyph dictionary for the purposes of extraction
# Instead of regular characters, it looks up their homoglyphs
reverse_unicode_dictionary = {value: key for key, value in unicode_dictionary.items()}

# Count non-whitespace lowercase characters in paragraphs
def count_chars_in_paragraphs(
        document: DocumentObject,
        index: int
    ) -> int:
    char_count = 0
    # Loop through paragraphs starting from the given index
    # And count all non-whitespace lowercase characters in the paragraph text using regex
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
    # Capacity is non-whitespace lowercase character amount in the document
    cap = char_count

    is_valid = stego_message_size_bits <= cap
    return is_valid

# Transform stego-message to bit string
def stego_message_to_bit_string(stego_message_bytes: bytes) -> str:
    stego_byte_to_8_bit_string = ''
    # For byte in stego-message
    # Transform it into an 8-bit string
    for byte in stego_message_bytes:
        stego_byte_to_8_bit_string += f"{byte:08b}"
    #print(len(stego_byte_to_8_bit_string))
    return stego_byte_to_8_bit_string

# Embedding algorithm
def embedding_in_run(
        run: Run,
        stego_message_text: str,
        stego_index: int,
        payload: int
    ) -> int:
    # Access the run's underlying XML element
    # To modify the XML structure of the run directly
    run_element = run._r
    # Either access the existing run properties 
    run_properties = run_element.rPr
    # Or create an empty run properties element
    # And insert it into the new run as its first child
    if run_properties == None:
        run_properties = OxmlElement("w:rPr")
        run_element.insert(0, run_properties)

    # The reasons for this is that
    # Since many homoglyphs are characters from other languages 
    # The proofing engine can draw attention to them via user-visible warnings etc.
    # It's best to explicitly tell the document to ignore any spelling errors produced by it
    # In order to minimize detection during steganalysis
    if run_properties.find(qn("w:noProof")) == None:
        run_properties.append(OxmlElement("w:noProof"))

    # Count all unused non-whitespace lowercase characters in the current run's text
    all_lowercase_letters = re.findall(r'[a-z]', run.text, flags=re.UNICODE)
    nr_of_unused_lowercase_letters = len(all_lowercase_letters)

    # Get the run's text, embedded text (initially empty), and remaining text (initially the full text)
    text = run.text
    embedded_text = ''
    remaining_text = text

    # Loop through all unused non-whitespace lowercase characters in the run
    for _ in range(nr_of_unused_lowercase_letters):
        # Check if the entire stego-message has already been embedded 
        if stego_index >= payload:
            break

        # Find the first non-whitespace lowercase character in the remaining text
        lowercase_letter = re.search(r'[a-z]', remaining_text, flags=re.UNICODE)
        # If no more non-whitespace lowercase characters are found, 
        # Add the remaining text to the embedded text
        # Empty the remaining text and exit the loop
        if lowercase_letter == None:
            embedded_text += remaining_text
            remaining_text = ''
            break
        # Otherwise
        else:
            # Get the index of the found non-whitespace lowercase character
            cover_symbol_index = lowercase_letter.start()
            # Get the found non-whitespace lowercase character
            single_cover_char = remaining_text[cover_symbol_index]
            # Add all remaining text before the found non-whitespace lowercase character as embedded text (left-side)
            embedded_text += remaining_text[:cover_symbol_index]
            # Update the remaining text to only include text after the found non-whitespace lowercase character (right-side)
            remaining_text = remaining_text[cover_symbol_index + 1:]
            # If the stego-bit string at the current index has 1 bit value
            if stego_message_text[stego_index] == '1':
                # Get the homoglyph for the stego character, default to the original character if not found (primarily for linter)
                homoglyph = unicode_dictionary.get(single_cover_char, single_cover_char)
                # And add said homoglyph to the embedded text as normal text
                embedded_text += homoglyph
            # Otherwise
            else:
                # Add the regular non-whitespace lowercase character
                embedded_text += single_cover_char
            # After embedding, increment the stego index
            stego_index += 1

    # If there is any remaining text left in the current run
    # Add it to the embedded text
    if remaining_text != '':
        embedded_text += remaining_text

    # If any embedding has occurred (embedded text differs from original text)
    # Update the run's text with the newly embedded text
    # That contains the original text and the embedded homoglyphs
    if embedded_text != text:
        run.text = embedded_text

    return stego_index

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    # Get the entire document text
    text = unified_stego_file.extract_text(document, True, 0)
    # Initialize an empty bytes object to hold the stego-message bytes
    stego_message_bytes = b''
    # Create a boolean to flag the detection of the starting homoglyph
    first_stego_char_found = False
    # Initialize the entire stego-character 8-bit string
    stego_8_bit_string = ''
    # And a completely 8-bit string
    end_8_bit_string = '00000000'
    # Loop through each character in the the document text
    for char in text:
        # Check if the character has a homoglyph value in the dictionary
        stego_char = reverse_unicode_dictionary.get(char)
        # While the first homoglyph has not been found
        if not first_stego_char_found:
            # If the gained character value from the dictionary is not empty
            if stego_char != None:
                # The starting homoglyph has been found
                first_stego_char_found = True
            # Otherwise keep searching
            continue
        # After the starting homoglyph has been found
        else:
            # If the character has a homoglpyh value in the dictionary
            if stego_char != None:
                # It represents 1 bit, add it to the stego-character 8-bit string
                stego_8_bit_string += '1'
            # If the character doesn't have a homoglpyh value in the dictionary
            # But is still a non-whitespace lowercase character
            elif 'a' <= char <= 'z':
                # It represents 0 bit, add it to the stego-character 8-bit string
                stego_8_bit_string += '0'
            # Otherwise the character holds no value, move on
            else:
                continue
            # If the stego-character 8-bit string has reached 8 bits
            if len(stego_8_bit_string) == 8:
                # If they are all 0s, the end of the stego-message has been reached
                if stego_8_bit_string == end_8_bit_string:
                    break
                # Otherwise
                else:
                    # Transform the 8-bit string into a byte
                    stego_byte = int(stego_8_bit_string, 2).to_bytes(1, 'big')
                    # Add the stego-byte to the stego-message bytes object  
                    stego_message_bytes += stego_byte
                    # And restart the 8-bit string
                    stego_8_bit_string = ''

    # Decode the stego-message from raw bytes into UTF-8 (readable text)
    stego_message = stego_message_bytes.decode('utf-8')
    # print(stego_message)
    return stego_message

#### Main function ###
def stego_method_unicode_homoglyphs() -> None:
    _, stego_message_bytes = unified_stego_file.stego_message()
    # Transform stego-message into a bit string
    # With one additional 1 bit to mark the beginning of the stego-message
    # For extraction purposes 
    stego_message_bytes_to_8_bit_string = '1' + stego_message_to_bit_string(stego_message_bytes)

    unified_stego_file.stego_method(
        'stego_method_6',
        (stego_message_bytes_to_8_bit_string, stego_message_bytes),
        count_chars_in_paragraphs,
        is_capacity_enough_for_message,
        embedding_in_run,
        'string',
        stego_message_extraction
    )

if __name__ == "__main__":
    stego_method_unicode_homoglyphs()