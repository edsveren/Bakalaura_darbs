import random
from pathlib import Path
from typing import Callable, Any
from docx import Document
from docx.text.run import Run
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn

# Get stego-message from file
def stego_message() -> tuple[str, bytes]:
    stego_message_text = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    stego_message_bytes = stego_message_text.encode("utf-8")
    return stego_message_text, stego_message_bytes

# Extract text from the document
def extract_text(
        document: DocumentObject,
        NBSP: bool,
        index: int
    ) -> str:
    text = []
    for paragraph in document.paragraphs[index:]:
        if NBSP == True:
            text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
        else:
            text.append(paragraph.text)
    text = "\n".join(text)
    text = ''.join(text)
    # print(f"Document text:\n{text}")
    return text

# Choose a random paragraph to embed the stego-message into
def choose_random_paragraph(
        document: DocumentObject,
        necessary_element_count_function: Callable[[DocumentObject, int], int],
        is_capacity_enough_for_message: Callable[[DocumentObject, int, int, int], bool],
        stego_message_size_bits: int
    ) -> int | None:

    paragraphs = document.paragraphs
    if not paragraphs:
        return None
    while True:
        random_paragraph_index = random.randint(0, len(paragraphs) - 1)
        random_paragraph = paragraphs[random_paragraph_index]
        necessary_element_count = necessary_element_count_function(document, random_paragraph_index)
        is_valid = is_capacity_enough_for_message(document, necessary_element_count, stego_message_size_bits, random_paragraph_index)
        if is_valid:
            print(f"Random paragraph start: {random_paragraph.text}")
            return random_paragraph_index

# Get stego-message bytes and bits     
def get_bytes_and_bits(stego_message_bytes: bytes) -> tuple[int, int]:
    stego_message_size_bytes = len(stego_message_bytes)
    stego_message_size_bits = 8 * stego_message_size_bytes
    
    # print(f"Stego-message bytes: {stego_message_size_bytes}")
    # print(f"Stego-message bites: {stego_message_size_bits}")

    return stego_message_size_bytes, stego_message_size_bits

### Main file ###   
def stego_method(
        stego_method: str,
        specialized_stego_message: tuple[str, bytes]|None,
        necessary_element_count_function: Callable[[DocumentObject, int], int],
        is_capacity_enough_for_message:  Callable[[DocumentObject, int, int, int], bool],
        embedding_algorithm: Callable[[Run, Any, int, int], int], # Any = str|bytes
        embedding_data_type: str,
        extraction_algorithm: Callable[[DocumentObject], str]
    ) -> None:
    # DOCX file
    base = "data_set/clean_files"
    for file in Path(base).iterdir():
        docPath = str(Path(f"{base}/{file.name}")) #Path("data_set/clean_files/TEST_0.docx")
        print(f"DOCX file: {docPath}")
        print("Beginning the embedding process...")
        document = Document(docPath)
        # text = extract_text(document, True, 0)
        necessary_element_count = necessary_element_count_function(document, 0)

        if specialized_stego_message == None:
            stego_message_text, stego_message_bytes = stego_message()
        else:
            stego_message_text, stego_message_bytes = specialized_stego_message
            
        stego_message_size_bytes, stego_message_size_bits = get_bytes_and_bits(stego_message_bytes)

        embedded = False
        while not embedded:
            # Check if the paragraph has enough runs to embed the message
            print("Checking if the cover object is valid for embedding...")
            is_valid = is_capacity_enough_for_message(document, necessary_element_count, stego_message_size_bits, 0)
            print("The cover object is valid:", is_valid)
            if not is_valid:
                print("Not enough capacity in the document to embed the message.")
                break

            # Embed stego-message in DOCX
            print("Embedding stego-message...")
            random_paragraph_index = choose_random_paragraph(document, necessary_element_count_function, is_capacity_enough_for_message, stego_message_size_bits)
            if random_paragraph_index is None:
                print("No paragraphs available for embedding.")
                break

            payload = stego_message_size_bytes
            stego_index = 0

            data_to_embed = ''
            if embedding_data_type == 'bytes':
                data_to_embed = stego_message_bytes
            elif embedding_data_type == 'string':
                data_to_embed = stego_message_text

            for paragraph in document.paragraphs [random_paragraph_index:]:
                if stego_index < payload:
                    original_run_amount = list(paragraph.runs)
                    for run in original_run_amount:
                        run_element = run._r
                        # Only process runs that contain text
                        if run_element.find(qn('w:t')) != None:
                            if stego_index < payload:
                                stego_index = embedding_algorithm(run, data_to_embed, stego_index, payload)
                                # next_stego_index = embedding_in_run(run, stego_message_text, stego_index, payload)
                                # stego_index = next_stego_index
                            else:
                                break
                else:
                    break
            
            #print("Extracting stego-message...")
            stego_message_text, _ = stego_message()
            extracted_stego_message = extraction_algorithm(document)
            print(f"Extracted stego-message: {extracted_stego_message}")
            if stego_message_text != extracted_stego_message:
                print("Extracted message is not equal to stego-message!")
                break
            #print("Extraction successful!")
            print("Embedding successful!")
            embedded = True

        if embedded:
            stegoDocPath = str(Path(f"data_set/stego_files/{stego_method}/{file.name}"))
            document.save(stegoDocPath)
            print(f"Saved: {stegoDocPath}")
        else:
            print("Embedding not possible.")
        print()