import re
from copy import deepcopy
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import scripts.stego_methods.unified_stego_file as unified_stego_file

# Count (black coloured) chars in paragraphs starting from index
def count_black_chars_in_paragraphs(
        document: DocumentObject,
        index: int
    ) -> int:
    black_char_count = 0
    for paragraph in document.paragraphs[index:]:
        for run in paragraph.runs:
            # Count all non-whitespace chars in runs
            # If the run's font color is black or automatic (default)
            text = ''
            run_properties = run._r.rPr
            if run_properties == None:
                text = run.text.replace('\xa0', '\x20')
            else:
                color_element = run_properties.find(qn('w:color'))
                # Run has no color property
                if color_element == None:
                    text = run.text.replace('\xa0', '\x20')
                else:
                    color_element_value = color_element.get(qn('w:val')) 
                    # Run color value is black or automatic (default)
                    if color_element_value == None or color_element_value.lower() in ('auto', '000000'):
                        text = run.text.replace('\xa0', '\x20')
                    else:
                        text = ''
            chars = re.findall(r'\S', text, flags=re.UNICODE)
            black_char_count += len(chars)
    # print(f"Total black coloured char count: {black_char_count}")
    return black_char_count

# Check if max capacity is enough for the message
def is_capacity_enough_for_message(
        document: DocumentObject,
        black_char_count: int,
        stego_message_size_bits: int,
        index: int
    ) -> bool:
    # # in theory
    # # Capacity is 8 bits per character
    # cap = 8 * char_count

    # in practice
    # Capacity is 8 bits per black coloured character
    cap = 8 * black_char_count

    is_valid = stego_message_size_bits <= cap
    return is_valid

# Create a new run
def insert_in_run(
        current_run: Run, # To modify the XML structure of the run directly
        base_run: Run, # To copy the run properties from
        char: str,
        stego_bits: str | None
    ) -> Run|None:

    current_run_element = current_run._r
    base_run_element = base_run._r
    new_run_element = OxmlElement('w:r')

    if base_run_element.rPr != None:
        new_run_element.append(deepcopy(base_run_element.rPr))
        new_run_properties = new_run_element.find(qn('w:rPr')) 
    else:
        new_run_properties = OxmlElement('w:rPr')
        new_run_element.insert(0, new_run_properties)

    new_text_element = OxmlElement('w:t')
    new_run_element.append(new_text_element)
    new_text_element.text = char
    
    if new_text_element.text.startswith('\x20') or new_text_element.text.endswith('\x20'):
        new_text_element.set(qn('xml:space'), 'preserve')

    # This stego-method depends on modifying the RGB color values of the run
    # If stego-bits are provided, modify the color property of the new run
    if stego_bits != None:
        color_element = new_run_properties.find(qn('w:color'))
        if color_element == None:
            color_element = OxmlElement('w:color')
            new_run_properties.append(color_element)
        color_element_value = color_element.get(qn('w:val'))
        # Only embed stego-bits if the run properties color value is black or automatic (default) 
        # The theory states to use any coloured text
        # But in practice, it's difficult to extract the stego-message
        # Without knowing a baseline, like black which has clean, consistent RGB(0,0,0) values
        # To do so would require some kind of stego-key
        if color_element_value == None or color_element_value.lower() in ('auto', '000000'):
            # Transform stego-bits into RGB values
            # Set 3 bits for Red, 3 bits for Green, 2 bits for Blue
            r_bit = int(stego_bits[0:3], 2)
            g_bit = int(stego_bits[3:6], 2)
            b_bit = int(stego_bits[6:8], 2)

            rgb_string = f"{r_bit:02x}{g_bit:02x}{b_bit:02x}"
            color_element.set(qn('w:val'), rgb_string)
        # If that's not possible, signal to the embedding algorithm
        # That the current run is exhausted for embedding
        # And avoid endlessly looping on it
        else:
            return None
    
    current_run_element.addnext(new_run_element)
    new_run = Run(new_run_element, current_run._parent)
    return new_run

# Splitting the existing runs into before, current and after
def slipt_run_for_embedding(
        run: Run,
        byte: int
    ) -> Run | None:

    # Find the first non-whitespace character in the run
    text = run.text
    cover_chars = re.search(r'\S', text, flags=re.UNICODE)
    # And if there are no non-whitespace characters, run is exhausted
    if cover_chars == None:
        return None

    cover_symbol_index = cover_chars.start()
    single_cover_char = text[cover_symbol_index]
    left_text = text[:cover_symbol_index]
    right_text = text[cover_symbol_index + 1:]

    stego_byte_to_binary_string = f"{byte:08b}"

    stego_char_run = insert_in_run(run, run, single_cover_char, stego_byte_to_binary_string)
    # If the run could not be created, the current run properties are not black etc.
    if stego_char_run == None:
        return None
    
    # Only deal with the current and remaining run text
    # If the stego-char run was created successfully
    # Otherwise, this would corrupt the document content and make it unusable
    if left_text != '':
        text_element = run._r.find(qn('w:t'))
        if text_element == None:
            text_element = OxmlElement('w:t')
            run._r.append(text_element)
        if left_text.startswith('\x20') or left_text.endswith('\x20'):
            text_element.set(qn('xml:space'), 'preserve')
        run.text = left_text

    remaining_run = insert_in_run(stego_char_run, run, right_text, None)

    # If left-side text is empty, cleanup the original run
    # Because empty runs are left as artifacts in the document
    # That can lead to easier discovery during steganalysis
    # Leaving only the stego-byte run and the remaining run
    if left_text == '':
        left_parent = run._r.getparent()
        left_parent.remove(run._r)
    return remaining_run

# Embedding algorithm
def embedding_in_run(
        run: Run,
        stego_message_bytes: bytes,
        stego_index: int,
        payload: int
    ) -> int:

    current_run = run
    non_whitespace_chars = re.findall(r'\S', run.text, flags=re.UNICODE)
    nr_of_non_whitespace_chars = len(non_whitespace_chars)

    for _ in range(nr_of_non_whitespace_chars):
        if stego_index < payload:
            next_run = slipt_run_for_embedding(current_run, stego_message_bytes[stego_index])
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
    stego_message_bytes = b''
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run_properties = run._r.rPr
            if run_properties != None:
                color_element = run_properties.find(qn('w:color'))
                if color_element != None:
                    color_element_value = color_element.get(qn('w:val'))
                    # In theory, any coloured text can be used
                    # But in practice, it requires some kind of stego-key
                    # Because it's difficult to extract the stego-message deterministically
                    # Without said key which provides original base values of non-stego-data
                    if color_element_value.lower() not in ('000000', 'auto'):
                        r_bit = int(color_element_value[0:2], 16)
                        g_bit = int(color_element_value[2:4], 16)
                        b_bit = int(color_element_value[4:6], 16)
                        # Red and Green color channels can contain max 111 (binary) = 7 (decimal) value
                        # Blue color channel can contain max 11 (binary) = 3 (decimal) value
                        if r_bit <= 7 and g_bit <= 7 and b_bit <= 3:
                            stego_bit_string = f"{r_bit:03b}{g_bit:03b}{b_bit:02b}"
                            stego_byte = int(stego_bit_string, 2).to_bytes(1, 'big')
                            # print(stego_byte)
                            stego_message_bytes += stego_byte
    # print(stego_message_bytes)
    stego_message = stego_message_bytes.decode('utf-8')
    # print(stego_message)
    return stego_message

### Main function ###
def stego_method_modify_RGB_color_ch() -> None:
    unified_stego_file.stego_method(
        'stego_method_4',
        None,
        count_black_chars_in_paragraphs,
        is_capacity_enough_for_message,
        embedding_in_run,
        'bytes',
        stego_message_extraction
    )

if __name__ == "__main__":
    stego_method_modify_RGB_color_ch()