import re
import random
from pathlib import Path
from docx import Document
from docx.text.run import Run
from docx.oxml.parser import OxmlElement
from docx.oxml.ns import qn
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

# Count words in paragraphs
def count_words_in_paragraphs(document: DocumentObject, index: int) -> int:
    word_count = 0
    for paragraph in document.paragraphs[index:]:
        text = paragraph.text.replace('\xa0', '\x20')  # NBSP -> space
        words = re.findall(r'\S+', text, flags=re.UNICODE)
        word_count += len(words)
    #print("Kopējais vārdu skaits:", word_count)
    return word_count

# Check if max capacity is enough for the message
def is_capacity_enough_for_message(word_count: int, stegoMessage_size_bits: int) -> bool:
    cap = 8 * (word_count - 1)
    is_valid = stegoMessage_size_bits <= cap
    return is_valid

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
def stego_message() -> tuple[str, bytes]:
    stegoMessageText = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    stegoMessage_bytes = stegoMessageText.encode("utf-8")
    #print("Stego-message data:", stegoMessageText)
    return stegoMessageText, stegoMessage_bytes

def stego_message_standarization_to_unispace_method(stegoMessageText: str) -> str:
    stegoMessageInWhiteSpaceUnicode = ''
    for line in stegoMessageText:
        for char in line:
            char = char.upper()
            key = reverse_unispace_dictionary.get(char)
            if key != None:
                stegoMessageInWhiteSpaceUnicode += char
    return stegoMessageInWhiteSpaceUnicode

# Choose random paragraph
def choose_random_paragraph(document: DocumentObject, stegoMessage_toBase64_size_bits: int) -> int | None:
    paragraphs = document.paragraphs
    if not paragraphs:
        return None
    while True:
        random_paragraph_index = random.randint(0, len(paragraphs) - 1)
        random_paragraph = paragraphs[random_paragraph_index]
        word_count = count_words_in_paragraphs(document, random_paragraph_index)
        is_valid = is_capacity_enough_for_message(word_count, stegoMessage_toBase64_size_bits)
        if is_valid:
            #print("Random paragraph start:", random_paragraph.text)
            return random_paragraph_index

# Embedding algorithm
def embedding_in_run(run: Run, stego_message_text: str, stego_index: int, payload: int) -> int:
    run_element = run._r
    text_whitespaces = re.findall(r'\x20', run.text, flags=re.UNICODE)
    nr_of_unused_whitespace = len(text_whitespaces)

    text = run.text
    embedded_text = ''
    remaining_text = text

    for _ in range(nr_of_unused_whitespace):
        if stego_index >= payload:
            break

        whitespace = re.search(r'\x20', remaining_text, flags=re.UNICODE)
        if whitespace is None:
            embedded_text += remaining_text
            remaining_text = ''
            break
        else:
            cover_whitespace_index = whitespace.start()
            embedded_text += remaining_text[:cover_whitespace_index]
            remaining_text = remaining_text[cover_whitespace_index + 1:]

            unispace_combination = reverse_unispace_dictionary.get(stego_message_text[stego_index])
            unispace_combination_string = ''.join(unispace_combination)
            embedded_text += unispace_combination_string
            stego_index += 1

    if remaining_text != '':
        embedded_text += remaining_text

    if embedded_text != text:
        run.text = embedded_text
        run_properties = run_element.rPr
        if run_properties is None:
            run_properties = OxmlElement("w:rPr")
            run_element.insert(0, run_properties)
        if run_properties.find(qn("w:noProof")) is None:
            run_properties.append(OxmlElement("w:noProof"))

        for text_element in run_element.iter(qn('w:t')):
            text_element.set(qn('xml:space'), 'preserve')

    return stego_index

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
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
    #print(stegoMessage)
    stegoMessage = stegoMessage[1:-1]
    return stegoMessage
            
### Main ###   
def main() -> None:          
    # DOCX file
    base = "data_set/clean_files"
    for file in Path(base).iterdir():
        docPath = str(Path(f"{base}/{file.name}")) #Path("data_set/clean_files/TEST_0.docx")
        print(f"DOCX file: {docPath}")
        print("Beginning the embedding process...")
        document = Document(docPath)
        text = extract_text(document)
        word_count = count_words_in_paragraphs(document, 0)

        stego_message_text, stegoMessage_bytes = stego_message()
        stegoMessage_size_bytes = len(stegoMessage_bytes)
        stegoMessage_size_bits = 8 * stegoMessage_size_bytes
        #print("Regular bytes:", stegoMessage_size_bytes)
        #print("Regular bites:", stegoMessage_size_bits)

        stegoMessageInWhiteSpaceUnicode = stego_message_standarization_to_unispace_method(stego_message_text)
        stegoMessageInWhiteSpaceUnicode_embedding_ready_format = '\x20' + stegoMessageInWhiteSpaceUnicode + '\x20'

        stegoMessageInWhiteSpaceUnicode_size_bytes = len(stegoMessageInWhiteSpaceUnicode)
        stegoMessageInWhiteSpaceUnicode_size_bits = 8 * stegoMessageInWhiteSpaceUnicode_size_bytes

        embedded = False
        while not embedded:
            # Check if the paragraph has enough runs to embed the message
            print("Checking if the cover object is valid for embedding...")
            is_valid = is_capacity_enough_for_message(word_count, stegoMessageInWhiteSpaceUnicode_size_bits)
            print("The cover object is valid:", is_valid)
            if not is_valid:
                print("Not enough capacity in the document to embed the message.")
                break

            random_paragraph_index = choose_random_paragraph(document, stegoMessageInWhiteSpaceUnicode_size_bits)
            if random_paragraph_index is None:
                print("No paragraphs available for embedding.")
                break

            # Embed stego-message in DOCX
            print("Embedding stego-message...")
            payload = stegoMessageInWhiteSpaceUnicode_size_bytes + 2
            stego_index = 0
            #while payload < stego_index:
            for paragraph in document.paragraphs [random_paragraph_index:]:
                if stego_index < payload:
                    original_run_amount = list(paragraph.runs)
                    for run in original_run_amount:
                        run_element = run._r
                        # Only process runs that contain text
                        if run_element.find(qn('w:t')) != None:
                            if stego_index < payload:
                                stego_index = embedding_in_run(run, stegoMessageInWhiteSpaceUnicode_embedding_ready_format, stego_index, payload)
                            else:
                                break
                else:
                    break

            #print("Extracting stego-message...")
            if stegoMessageInWhiteSpaceUnicode != stego_message_extraction(document):
                print("Extracted message is not equal to stego-message!")
                break
            #print("Extraction successful!")
            print("Embedding successful!")   
            embedded = True

        if embedded:
            stegoDocPath = str(Path(f"data_set/stego_files/stego_method_5/{file.name}"))
            document.save(stegoDocPath)
            print(f"Saved: {stegoDocPath}")
        else:
            print("Embedding not possible.")
        print()

if __name__ == "__main__":
    main()