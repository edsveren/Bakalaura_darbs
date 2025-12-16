import unified_passive_attack_file
from docx.document import Document as DocumentObject

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

# Transform stego-message to bit string
# def stego_message_to_bit_string(stegoMessage_bytes: bytes) -> str:
#     stego_byte_to_binary_string = ''
#     for byte in stegoMessage_bytes:
#         stego_byte_to_binary_string += f"{byte:08b}"
#     #print(len(stego_byte_to_binary_string))
#     return stego_byte_to_binary_string

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    # Get the entire document text
    text = unified_passive_attack_file.extract_text(document)
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

    try:
        #print(stegoMessage_as_base64)
        # Decode the stego-message from raw bytes into UTF-8 (readable text)
        stego_message = stego_message_bytes.decode('utf-8')
    except Exception as e: 
        print(e)
        stego_message = "TOO CORRUPT"
    #print(stegoMessage)
    return stego_message

def main(desired_file: str|None) -> None:
    unified_passive_attack_file.passive_attack("stego_method_6", stego_message_extraction, None, desired_file)
            
if __name__ == "__main__":
    main(None)
