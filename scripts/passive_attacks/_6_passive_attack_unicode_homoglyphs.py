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
    text = unified_passive_attack_file.extract_text(document)
    stegoMessage_bytes = b''
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
                    stegoMessage_bytes += stego_byte
                    stego_byte_string = ''
    try:
        #print(stegoMessage_as_base64)
        stegoMessage = stegoMessage_bytes.decode('utf-8', errors='replace')
    except Exception as e: 
        print(e)
        stegoMessage = "CORRUPTED"
    #print(stegoMessage)
    return stegoMessage

if __name__ == "__main__":
    unified_passive_attack_file.passive_attack("stego_method_6", stego_message_extraction, None)
