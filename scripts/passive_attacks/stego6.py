from pathlib import Path
from docx import Document

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

# Extract text from the document
def extract_text(document) -> str:
    text = []
    for paragraph in document.paragraphs:
        text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
    text = "\n".join(text)
    text = ''.join(text)
    #print("Teksts no dokumenta:\n", text)
    return text

# Stego-message
def stego_message() -> list[str]:
    stegoMessageText = Path("stego-messages/stego-message.txt").read_text(encoding="utf-8")
    #print("Stego-message data:", stegoMessageText)
    return stegoMessageText

# Transform stego-message to bit string
def stego_message_to_bit_string(stegoMessage_bytes) -> str:
    stego_byte_to_binary_string = ''
    for byte in stegoMessage_bytes:
        stego_byte_to_binary_string += f"{byte:08b}"
    #print(len(stego_byte_to_binary_string))
    return stego_byte_to_binary_string

# Extraction algorithm
def stego_message_extraction(document) -> str:
    text = extract_text(document)
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
    stegoMessage = stegoMessage_bytes.decode('utf-8')
    #print(stegoMessage)
    return stegoMessage

# DOCX file
docPath = Path("data_set/attacked_stego-files/stego-method_1/TEST_0.docx")
document = Document(docPath)
text = extract_text(document)

stego_message_text = stego_message()

### Main

print("Extracting stego-message...")
stego_message_extracted = stego_message_extraction(document)
if stego_message_extracted != '':
    print(f"The extracted stego-message: {stego_message_extracted}")
    if stego_message_text == stego_message_extracted:
        print("Extraction successful!")
    else:
        print("Extracted message is not equal to stego-message!")
else:
    print(f"The extracted stego-message: {None}")