from pathlib import Path
from collections import Counter
from docx import Document
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

# Extract text from the document
def extract_text(document: DocumentObject) -> str:
    text = []
    for paragraph in document.paragraphs:
        text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
    text = "\n".join(text)
    text = ''.join(text)
    #print("Teksts no dokumenta:\n", text)
    return text

# Stego-message
def stego_message() -> str:
    stegoMessageText = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    #print("Stego-message data:", stegoMessageText)
    return stegoMessageText

# Transform stego-message to bit string
def stego_message_to_bit_string(stegoMessage_bytes: bytes) -> str:
    stego_byte_to_binary_string = ''
    for byte in stegoMessage_bytes:
        stego_byte_to_binary_string += f"{byte:08b}"
    #print(len(stego_byte_to_binary_string))
    return stego_byte_to_binary_string

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
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
    try:
        #print(stegoMessage_as_base64)
        stegoMessage = stegoMessage_bytes.decode('utf-8', errors='replace')
    except Exception as e: 
        print(e)
        stegoMessage = "CORRUPTED"
    #print(stegoMessage)
    return stegoMessage

def check_for_stego_message(file_name: str, document: DocumentObject, stego_message_text: str) -> str:
    stego_message_extracted = stego_message_extraction(document)
    if stego_message_extracted != '':
        #print(stego_message_extracted)
        if stego_message_text == stego_message_extracted:
            #print(f"{file_name}'s extracted stego-message: {stego_message_extracted}. EQUAL!")
            return "SAFE" # STEGO-MESSAGE GOOD
        else:
            #print(f"{file_name}'s extracted stego-message: {stego_message_extracted}. NOT EQUAL!")
            return "CORRUPTED" # STAGE 2: STEGO-MESSAGE DEGRADED
    else:
        #print(f"{file_name}'s extracted stego-message: THERE IS NO STEGO-MESSAGE!")
        return "MISSING" # THERE IS NO STEGO-MESSAGE

### Main ###
def main() -> None:
    # DOCX file
    attacked_stego_files = "data_set/attacked_stego_files"
    stego_message_text = stego_message()
    print()
    for attack_directories in Path(attacked_stego_files).iterdir():
        print(attack_directories.name)
        states_list = []
        nr_of_files = 0
        print("Extracting stego-message...")
        for stego_directories in attack_directories.iterdir():
            #print(stego_directories.name)
            if stego_directories.name == "stego_method_6":
                for file in stego_directories.iterdir():
                    #docPath = f"{attacked_stego_files}/{stego_directories.name}/{file.name}"
                    if file.is_file() and not file.name.startswith("."):
                        docPath = str(Path(f"{attacked_stego_files}/{attack_directories.name}/{stego_directories.name}/{file.name}"))
                        document = Document(docPath)
                        #text = extract_text(document)
                        #print(file.name)
                        state = check_for_stego_message(file.name, document, stego_message_text)
                        states_list.append(state)
                        nr_of_files += 1
        #break
        counter = {key: Counter(states_list).get(key, 0) for key in ['SAFE', 'CORRUPTED', 'MISSING']}
        # sizes = list(counter.keys())
        # frequencies = list(counter.values())
        # print(sizes)
        # print(frequencies)
        for size, frequency in counter.items():
            print(f"State: {size}, Amount: {frequency} out of {nr_of_files}")
        print()
    print("Extraction over!")
            
if __name__ == "__main__":
    main()