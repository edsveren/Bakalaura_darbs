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
    # Loop through paragraphs starting from the given index
    # And count all words in the paragraph text using regex
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
    # # in theory
    # equal_mark = "="
    # equal_mark_count = 0
    # # Loop through the stego-message in Base64 to count '=' characters (ought to be 2)
    # for char in stego_message_to_Base64_text:
    #    if char == equal_mark:
    #        equal_mark_count += 1
    # # Capacity is 6 bits per whitespace in-between words, minus the '=' characters
    # # Which do not carry any information
    # cap = 6 * (word_count - 1 - equal_mark_count)

    # in practice
    # Get the document text from the beginning or from the random starting index
    text = unified_stego_file.extract_text(document, False, index)
    # Count all whitespaces in the text using regex
    text_whitespaces = re.findall(r'\x20', text, flags=re.UNICODE)
    # With each whitespace able to carry 6 bits of Base64 information
    cap = 6 * len(text_whitespaces)

    # The amount of stego-message bits must be less than or equal to the capacity
    is_valid = stego_message_size_bits <= cap
    return is_valid

# Encode stego-message in Base64 format
def stego_message_base64(stego_message_bytes: bytes) -> tuple[str, bytes]:
    stego_message_to_Base64_bytes = base64.b64encode(stego_message_bytes)
    stego_message_to_Base64_text = stego_message_to_Base64_bytes.decode('ascii')
    # print(f"Stego-message in Base64: {stego_message_to_Base64_text}")
    return stego_message_to_Base64_text, stego_message_to_Base64_bytes

# Create a new run
def insert_in_run(
        current_run: Run, # To modify the XML structure of the run directly
        base_run: Run, # To copy the run properties from
        char: str,
        type: str | None
    ) -> Run:

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
    # Based on type (whitespace, stego-character or simply the remaining text)
    match type:
        case 'whitespace':
            new_text_element.text = char
            # To avoid XML parsers from stripping whitespace characters
            # Automatically set the space preservation attribute on the text element
            new_text_element.set(qn('xml:space'), 'preserve')
        case 'stego_char':
            new_text_element.text = char

            # This stego-method depends on modifying the font color, size and visibility (hidden function)
            # So these run properties are set specifically for the stego-character run
            # All three properties are either modified or created anew if not present

            # Run color gets set to white (FFFFFF)
            color_element = new_run_properties.find(qn('w:color'))
            if color_element == None:
                color_element = OxmlElement('w:color')
                new_run_properties.append(color_element)
            color_element.set(qn('w:val'), 'FFFFFF')

            # Run font size gets set to 1px (2 half-points)
            font_size_element = new_run_properties.find(qn('w:sz'))
            if font_size_element == None:
                font_size_element = OxmlElement('w:sz')
                new_run_properties.append(font_size_element)
            font_size_element.set(qn('w:val'), '2')

            # Hidden function (vanish) gets enabled
            vanish_element = new_run_properties.find(qn('w:vanish'))
            if vanish_element == None:
                vanish_element = OxmlElement('w:vanish')
                new_run_properties.append(vanish_element)
            vanish_element.set(qn('w:val'), 'true')
        case _:
            new_text_element.text = char           
    
    # Add the new run element directly after the current run element in the Document XML tree
    current_run_element.addnext(new_run_element)
    # Wrap the new run element in a Run object API class
    # So that it can be manipulated for further management
    new_run = Run(new_run_element, current_run._parent)
    return new_run

# Splitting the existing runs into
# Before the insertion, inserted data and after insertion
def slipt_run_for_embedding(
        run: Run,
        stego_char: str
    ) -> Run | None:

    # Find the first whitespace in the run text
    text = run.text
    whitespace = text.find('\x20')
    # And if there are no whitespaces, run is exhausted
    if whitespace == -1:
        return None

    # Create text before the first whitespace (left)
    left_text = text[:whitespace] 
    # Create text after the first whitespace (right)
    right_text = text[whitespace + 1:]

    # Replace current run text with left-side text
    run.text = left_text

    # Create 4 new run elements for each part: left whitespace, stego-character, right whitespace, remaining text
    # Insert whitespace left of the stego-character in a new run
    left_whitespace = insert_in_run(run, run, '\x20', 'whitespace')

    # Insert the stego-character in a new run
    stego_char_run = insert_in_run(left_whitespace, run, stego_char, 'stego_char')

    # Insert whitespace right of the stego-character in a new run
    right_whitespace = insert_in_run(stego_char_run, run, '\x20', 'whitespace')

    # Replace the remaining run text with the right-side text
    remaining_run = insert_in_run(right_whitespace, run, right_text, None)
    return remaining_run

# Embedding algorithm
def embedding_in_run(
        run: Run,
        stego_message_to_Base64_text: str,
        stego_index: int,
        payload: int
    ) -> int:
    # Get the current run and count all unused whitespaces in its text
    current_run = run
    text_whitespaces = re.findall(r'\x20', run.text, flags=re.UNICODE)
    nr_of_unused_whitespace = len(text_whitespaces)

    # For each unused whitespace, embed the next stego-character from the stego-message in Base64
    # Increasing the stego_index until the run is exhausted or the payload is reached
    for _ in range(nr_of_unused_whitespace):
        if stego_index < payload:
            next_run = slipt_run_for_embedding(current_run, stego_message_to_Base64_text[stego_index])
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
    # Initialize an empty string to hold the stego-message in Base64
    stego_message_as_base64 = ''
    # Loop through all document paragraphs
    for paragraph in document.paragraphs:
        # Loop through all runs in the paragraph
        for run in paragraph.runs:
            # Get the run element's properties
            run_properties = run._r.rPr
            # Only analyze runs with run properties
            if run_properties != None:
                # Only analyze runs with exactly one character
                if len(run.text) == 1:
                    # Find the run property elements for color, size and hidden function
                    color_element = run_properties.find(qn('w:color'))
                    font_size_element = run_properties.find(qn('w:sz'))
                    vanish_element = run_properties.find(qn('w:vanish'))

                    # Check if all three run property elements are present
                    if None not in (color_element, font_size_element, vanish_element):
                        # Get their values
                        color_element_value = color_element.get(qn('w:val'))
                        font_size_value = font_size_element.get(qn('w:val'))
                        # And check if said values match the stego-embedding criteria
                        # Color = white (FFFFFF), size = 2 (1px), hidden = true (exists)
                        if color_element_value.upper() == 'FFFFFF' and font_size_value == '2':
                            # The run contains a stego-character
                            # Append it to the stego-message in Base64
                            stego_message_as_base64 += run.text

    # print(stego_message_as_base64)
    # Decode the stego-message from Base64 into UTF-8 (readable text)
    stego_message = base64.b64decode(stego_message_as_base64).decode('utf-8')
    # print(stego_message)
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