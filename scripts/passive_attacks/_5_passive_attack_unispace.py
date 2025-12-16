import re
import unified_passive_attack_file
from docx.document import Document as DocumentObject

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

def stego_message_standarization_to_unispace_method(stegoMessageText: str) -> str:
    stegoMessageInWhiteSpaceUnicode = ''
    for line in stegoMessageText:
        for char in line:
            char = char.upper()
            key = reverse_unispace_dictionary.get(char)
            if key != None:
                stegoMessageInWhiteSpaceUnicode += char
    return stegoMessageInWhiteSpaceUnicode

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    # Get the entire document text
    text = unified_passive_attack_file.extract_text(document)
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

def main(desired_file: str|None) -> None:
    stego_message_text = unified_passive_attack_file.stego_message()
    stegoMessageInWhiteSpaceUnicode = stego_message_standarization_to_unispace_method(stego_message_text)
    unified_passive_attack_file.passive_attack("stego_method_5", stego_message_extraction, stegoMessageInWhiteSpaceUnicode, desired_file)
            
if __name__ == "__main__":
    main(None)