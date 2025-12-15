import re
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject
import unified_stego_file

zero = '\u2009' # Thin space
one = '\u200A'  # Hair space
two = '\u200B'  # Zero width space

# Unispace dictionary mapping unispace combinations (tuples) to characters
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
    (two, two, two): '\x20', # whitespace
}

# Reverse unispace dictionary for the purposes of embedding
# Instead of combinations, it looks up characters to get the combination tuple
reverse_unispace_dictionary = {value: key for key, value in unispace_dictionary.items()}

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
    # Capacity is 8 bits per whitespace between words
    # If the 8-bit (1 byte) represents an uppercase ASCII character
    cap = 8 * (word_count - 1)

    is_valid = stego_message_size_bits <= cap
    return is_valid

# Standardize stego-message to unispace method
def stego_message_standarization_to_unispace_method(stego_message_text: str) -> str:
    # Initialize empty string to hold stego-message in unispace format
    stego_message_in_whitespace_unicode = ''
    
    # Loop each line in the text file
    for line in stego_message_text:
        # Loop each character in the line
        for char in line:
            # Transform character to uppercase
            char = char.upper()
            # Get the corresponding unispace combination for said character
            key = reverse_unispace_dictionary.get(char)
            # If the character is a valid value in the unispace dictionary,
            # Add it to the stego-message
            # Characters that are not in the dictionary
            # Are effectively removed from the stego-message
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

    # Count all unused whitespaces in the current run's text
    text_whitespaces = re.findall(r'\x20', run.text, flags=re.UNICODE)
    nr_of_unused_whitespace = len(text_whitespaces)

    # Get the run's text, embedded text (initially empty), and remaining text (initially the full text)
    text = run.text
    embedded_text = ''
    remaining_text = text

    # Loop through all unused whitespaces in the run
    for _ in range(nr_of_unused_whitespace):
        # Check if the entire stego-message has already been embedded 
        if stego_index >= payload:
            break

        # Find the first whitespace in the remaining text
        whitespace = re.search(r'\x20', remaining_text, flags=re.UNICODE)
        # If no more whitespaces are found, 
        # Add the remaining text to the embedded text
        # Empty the remaining text and exit the loop
        if whitespace == None:
            embedded_text += remaining_text
            remaining_text = ''
            break
        # Otherwise
        else:
            # Get the index of the found whitespace
            cover_whitespace_index = whitespace.start()
            # Add all remaining text before the found whitespace as embedded text (left-side)
            embedded_text += remaining_text[:cover_whitespace_index]
            # Update the remaining text to only include text after the found whitespace (right-side)
            remaining_text = remaining_text[cover_whitespace_index + 1:]

            # Get the unispace combination for the stego character, default to empty tuple if not found (primarily for linter)
            unispace_combination = reverse_unispace_dictionary.get(stego_message_text[stego_index], ())
            # Convert the unispace combination tuple into usable string
            unispace_combination_string = ''.join(unispace_combination)
            # And add said string to the embedded text as normal text
            embedded_text += unispace_combination_string
            # After embedding, increment the stego index
            stego_index += 1

    # If there is any remaining text left in the current run
    # Add it to the embedded text
    if remaining_text != '':
        embedded_text += remaining_text
    
    # Access the run's underlying XML element
    # To modify the XML structure of the run directly
    run_element = run._r

    # If any embedding has occurred (embedded text differs from original text)
    if embedded_text != text:
        # Update the run's text with the newly embedded text
        # That contains the original text and the embedded unispace characters
        run.text = embedded_text
        # Since unispace characters are whitespace characters
        # The XML structure needs to be modified to preserve whitespace
        # Otherwise, the XML parser may collapse or ignore consecutive whitespace characters
        # Either access the existing run properties 
        run_properties = run_element.rPr
        # Or create an empty run properties element
        # And insert it into the new run as its first child
        if run_properties == None:
            run_properties = OxmlElement("w:rPr")
            run_element.insert(0, run_properties)
            
        # Since the unispace method's whitespace characters can potentially cause confusion 
        # To the proofing engine, which can then draw attention to the whitespaces via user-visible warnings etc.
        # It's best to explicitly tell the document to ignore any spelling errors produced by it
        # In order to minimize detection during steganalyis
        if run_properties.find(qn("w:noProof")) == None:
            run_properties.append(OxmlElement("w:noProof"))
    
        # To avoid XML parsers from stripping whitespace characters which store the stego-data
        # Space preservation attribute is set to all text elements in the run element
        for text_element in run_element.iter(qn('w:t')):
            text_element.set(qn('xml:space'), 'preserve')

    return stego_index

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    # Get the entire document text
    text = unified_stego_file.extract_text(document, True, 0)
    # Create a dictionary where all unispace dictionary combination tuples
    # Are turned into 3 characters strings
    unispace_combination_as_string_dictionary = {''.join(key): value for key, value in unispace_dictionary.items()}
    # Initialize an empty string to hold the stego-message
    stego_message = ''
    # Search for the first starting unispace character combination
    zero_width_space = two * 3
    first_zero_width_space = re.search(zero_width_space, text, flags=re.UNICODE)
    # If the first starting unispace character combination doesn't exist, stego-message is missing
    if first_zero_width_space == None:
        return ''
    # Mark the beginning of the cover text containing the stego-message
    cover_text_with_stego = text[first_zero_width_space.start():]

    # Find the first regular whitespace starting from the first unispace character combination
    first_whitespace = re.search(r'\x20', cover_text_with_stego, flags=re.UNICODE)
    # If there is a regular whitespace, 
    # Use it to mark the end of the cover text containing the stego-message
    if first_whitespace != None:
        cover_text_with_stego = cover_text_with_stego[:first_whitespace.start()]

    # Initialize the tracking of individual 3 whitespace combinations
    unispace_combination_length = 0
    # And the entire stego-character unispace combination string
    unispace_combination = ''
    # Loop through the cover text
    for char in cover_text_with_stego:
        # If a char is any of the unispace whitespace characters
        if char in (zero, one, two):
            # Add that whitespace character to the unispace combination string
            unispace_combination += char
            # And increase the tracker's length
            unispace_combination_length += 1
            # Once the tracker has collected 3 whitespaces
            # It means one complete unispace combination
            # Containing a stego-character has been found
            if unispace_combination_length == 3:
                # Transform the unispace string into a stego-character
                stego_char = unispace_combination_as_string_dictionary.get(unispace_combination)
                # There shouldn't be an incorrect transformation
                # But just in case check if the stego-character is not None
                if stego_char != None:
                    # After successful transformation,
                    # Add the stego-character to the stego-message
                    # And restart the unispace combination length tracker and string
                    stego_message += stego_char
                    unispace_combination_length = 0
                    unispace_combination = ''
    # print(stego_message)
    # Remove the first and last boundary whitespaces
    stego_message = stego_message[1:-1]
    return stego_message

def stego_method_unispace() -> None:
    stego_message_text, stego_message_bytes = unified_stego_file.stego_message()
    # Standardize stego-message to a usable format for the unispace method
    stego_message_in_whitespace_unicode = stego_message_standarization_to_unispace_method(stego_message_text)
    # Add whitespace characters at the start and end of the standardized stego-message
    # Which designate the stego-message's beginning and end
    # For the purposes of extraction
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