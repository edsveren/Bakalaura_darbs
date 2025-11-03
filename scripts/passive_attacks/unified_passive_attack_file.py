from pathlib import Path
from collections import Counter
from docx import Document
from docx.document import Document as DocumentObject

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

def passive_attack(stego_method: str) -> None:
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
            if stego_directories.name == stego_method:
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