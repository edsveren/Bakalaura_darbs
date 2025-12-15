import re
from copy import deepcopy
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import unified_stego_file

# Count (black coloured) chars in paragraphs starting from index
def count_black_chars_in_paragraphs(
        document: DocumentObject,
        index: int
    ) -> int:
    black_char_count = 0
    # Loop through paragraphs starting from the given index
    for paragraph in document.paragraphs[index:]:
        # Loop through runs in the paragraph
        for run in paragraph.runs:
            # Count all non-whitespace chars in runs
            # If the run's font color is black or automatic (default)
            text = ''
            run_properties = run._r.rPr
            # Run has no properties
            if run_properties == None:
                text = run.text.replace('\xa0', '\x20')  # NBSP -> space
            else:
                color_element = run_properties.find(qn('w:color'))
                # Run has no color property
                if color_element == None:
                    text = run.text.replace('\xa0', '\x20')  # NBSP -> space
                else:
                    color_element_value = color_element.get(qn('w:val')) 
                    # Run color value is black or automatic (default)
                    if color_element_value == None or color_element_value.lower() in ('auto', '000000'):
                        text = run.text.replace('\xa0', '\x20')  # NBSP -> space
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

    # To modify the XML structure of the run directly
    # The underlying XML element of the run must be accessed
    current_run_element = current_run._r
    base_run_element = base_run._r

    # Create a new run element that will become the current run's next sibling
    new_run_element = OxmlElement('w:r')

    # Create new run element's properties
    if base_run_element.rPr != None:
        # Copy all existing run properties to the new run using deepcopy
        # Which recursively copies all nested child elements
        # And leaves no references to the original elements
        new_run_element.append(deepcopy(base_run_element.rPr))
        new_run_properties = new_run_element.find(qn('w:rPr')) 
    else:
        # Otherwise create an empty run properties element
        # And insert it into the new run as its first child
        new_run_properties = OxmlElement('w:rPr')
        new_run_element.insert(0, new_run_properties)

    # Create a new text element for the new run
    # And append it to the new run element
    new_text_element = OxmlElement('w:t')
    new_run_element.append(new_text_element)

    # Insert character in the new text element
    new_text_element.text = char
    
    # If the char is whitespace or otherwise starts or ends with a whitespace, preserve it
    if new_text_element.text.startswith('\x20') or new_text_element.text.endswith('\x20'):
        new_text_element.set(qn('xml:space'), 'preserve')

    # This stego-method depends on modifying the RGB color values of the run
    # If stego bits are provided, modify the color property of the new run
    if stego_bits != None:
        # Either modify or create the color element in the new run properties
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
        # If that's not possible, exit completely
        else:
            return None
    
    # Add the new run element directly after the current run element in the Document XML tree
    current_run_element.addnext(new_run_element)
    # Wrap the new run element in a Run object API class
    # So that it can be manipulated for further management
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

    # Get the first non-whitespace character index
    cover_symbol_index = cover_chars.start()
    # Get the first non-whitespace character
    single_cover_char = text[cover_symbol_index]
    # Create text before the first non-whitespace char (left) 
    left_text = text[:cover_symbol_index]
    # Create text after the first non-whitespace char (right)
    right_text = text[cover_symbol_index + 1:]

    # Transform stego-byte into 8-bit string
    stego_byte_to_binary_string = f"{byte:08b}"

    # Create a new run for the stego-byte
    stego_char_run = insert_in_run(run, run, single_cover_char, stego_byte_to_binary_string)
    # And if the run could not be created,
    # It means that its color properties were not suitable for embedding
    if stego_char_run == None:
        return None
    
    # Only deal with the current and remaining run text
    # If the stego-char run was created successfully
    # Otherwise, this would corrupt the document content and make it unusable
    
    # Only process the left-side text if there remains any
    # Otherwise, this would create empty runs and leave artifacts in the document
    # That can lead to easy discovery during steganalysis
    if left_text != '':
        # Ensure the text element exists
        # To be able to set xml:space property if needed
        text_element = run._r.find(qn('w:t'))
        if text_element == None:
            text_element = OxmlElement('w:t')
            run._r.append(text_element)
        # If the left-side text starts or ends with a whitespace, preserve it
        if left_text.startswith(' ') or left_text.endswith(' '):
            text_element.set(qn('xml:space'), 'preserve')
        # Replace current run text with left-side text
        run.text = left_text

    # Replace the remaining run text with the right-side text
    remaining_run = insert_in_run(stego_char_run, run, right_text, None)

    # If left-side text is empty, cleanup the original run
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

    # Get the current run and count all unused non-whitespace chars in its text
    current_run = run
    non_whitespace_chars = re.findall(r'\S', run.text, flags=re.UNICODE)
    nr_of_non_whitespace_chars = len(non_whitespace_chars)

    # For each unused non-whitespace char, embed the next stego byte from the stego-message
    # Increasing the stego_index until the run is exhausted or the payload is reached
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
    # Initialize an empty bytes object to hold the stego-message bytes
    stego_message_bytes = b''
    # Loop through all document paragraphs
    for paragraph in document.paragraphs:
        # Loop through all runs in the paragraph
        for run in paragraph.runs:
            # Get the run element's properties
            run_properties = run._r.rPr
            # Only analyze runs with run properties
            if run_properties != None:
                # Find the color element in the run properties
                color_element = run_properties.find(qn('w:color'))
                # And check if it is present
                if color_element != None:
                    # Get its value
                    color_element_value = color_element.get(qn('w:val'))
                    # And check if it is not black or automatic (default)
                    # In theory, any coloured text can be used
                    # But in practice, it requires some kind of stego-key
                    # Because it's difficult to extract the stego-message deterministically
                    # Without said key which provides original base values of non-stego data
                    if color_element_value.lower() not in ('000000', 'auto'):
                        # Get each color channel's bits from the RGB hex value
                        r_bit = int(color_element_value[0:2], 16)
                        g_bit = int(color_element_value[2:4], 16)
                        b_bit = int(color_element_value[4:6], 16)
                        # Red and Green color channels can contain max 111 (binary) = 7 (decimal) value
                        # Blue color channel can contain max 11 (binary) = 3 (decimal) value
                        if r_bit <= 7 and g_bit <= 7 and b_bit <= 3:
                            # Combine the bits into a single 8-bit string
                            stego_bit_string = f"{r_bit:03b}{g_bit:03b}{b_bit:02b}"
                            # And transform it into a byte
                            stego_byte = int(stego_bit_string, 2).to_bytes(1, 'big')
                            # print(stego_byte)
                            # Add the stego-byte to the stego-message bytes object
                            stego_message_bytes += stego_byte
    # print(stego_message_bytes)
    # Decode the stego-message from raw bytes into UTF-8 (readable text)
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