import re
from pathlib import Path
from docx import Document

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

# Extract text from the document
def extract_text(document) -> str:
    text = []
    for paragraph in document.paragraphs:
        text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
    text = "\n".join(text)
    #print("Teksts no dokumenta:\n", text)
    return text

# Stego-message
def stego_message() -> list[str]:
    stegoMessageText = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    #print("Stego-message data:", stegoMessageText)
    return stegoMessageText

def stego_message_standarization_to_unispace_method(stegoMessageText) -> str:
    stegoMessageInWhiteSpaceUnicode = ''
    for line in stegoMessageText:
        for char in line:
            char = char.upper()
            key = reverse_unispace_dictionary.get(char)
            if key != None:
                stegoMessageInWhiteSpaceUnicode += char
    return stegoMessageInWhiteSpaceUnicode

# Extraction algorithm
def stego_message_extraction(document) -> str:
    text = extract_text(document)
    unispace_combination_as_string_dictionary = {''.join(key): value for key, value in unispace_dictionary.items()}
    stegoMessage = ''
    first_zero_width_space = re.search(two, text, flags=re.UNICODE)
    if first_zero_width_space == None:
        return ''
    cover_text_with_stego = text[first_zero_width_space.start():]

    first_whitespace = re.search(r'\x20', cover_text_with_stego, flags=re.UNICODE)
    cover_text_with_stego = cover_text_with_stego[:first_whitespace.start()]

    unispace_combination_length = 0
    unispace_combination = ''
    for char in cover_text_with_stego:
        if char in (zero, one, two):
            unispace_combination += char
            unispace_combination_length += 1
            if unispace_combination_length == 3:
                stego_char = unispace_combination_as_string_dictionary.get(unispace_combination)
                if stego_char != None:
                    stegoMessage += stego_char
                    unispace_combination_length = 0
                    unispace_combination = ''
    return stegoMessage
            
# DOCX file
docPath = Path("data_set/attacked_stego_files/stego_method_1/TEST_0.docx")
document = Document(docPath)
text = extract_text(document)

stego_message_text = stego_message()
stegoMessageInWhiteSpaceUnicode = stego_message_standarization_to_unispace_method(stego_message_text)

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