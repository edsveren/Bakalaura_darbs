import random
from pathlib import Path
from typing import Callable
from docx import Document
from docx.text.run import Run
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn

# Stego-message
def stego_message() -> tuple[str, bytes]:
    stegoMessageText = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    stegoMessage_bytes = stegoMessageText.encode("utf-8")
    return stegoMessageText, stegoMessage_bytes

# Choose random paragraph
def choose_random_paragraph(
        document: DocumentObject,
        necessary_element_count_function: Callable[[DocumentObject, int], int],
        is_capacity_enough_for_message: Callable[[DocumentObject, int, int|None], bool],
        stegoMessage_size_bits: int
    ) -> int | None:

    paragraphs = document.paragraphs
    if not paragraphs:
        return None
    while True:
        random_paragraph_index = random.randint(0, len(paragraphs) - 1)
        random_paragraph = paragraphs[random_paragraph_index]
        necessary_element_count = necessary_element_count_function(document, random_paragraph_index)
        is_valid = is_capacity_enough_for_message(document, stegoMessage_size_bits, random_paragraph_index)
        if is_valid:
            #print("Random paragraph start:", random_paragraph.text)
            return random_paragraph_index

### Main ###   
def main(
        stego_method: str,
        embedding_in_run: Callable[[Run, str, int, int], int],
        is_capacity_enough_for_message:  Callable[[DocumentObject, int, int|None], bool],
        stego_message_extraction: Callable[[DocumentObject], str],
    ) -> None:
    # DOCX file
    base = "data_set/clean_files"
    for file in Path(base).iterdir():
        docPath = str(Path(f"{base}/{file.name}")) #Path("data_set/clean_files/TEST_0.docx")
        print(f"DOCX file: {docPath}")
        print("Beginning the embedding process...")
        document = Document(docPath)
        # text = extract_text(document, True)
        # word_count = count_words_in_paragraphs(document, 0)

        stego_message_text, stegoMessage_bytes = stego_message()
        stegoMessage_size_bytes = len(stegoMessage_bytes)
        stegoMessage_size_bits = 8 * stegoMessage_size_bytes
        #print("Regular bytes:", stegoMessage_size_bytes)
        #print("Regular bites:", stegoMessage_size_bits)

        stegoMessage_toBase64_text, stegoMessage_toBase64_bytes = stego_message_base64(stegoMessage_bytes)
        stegoMessage_toBase64_size_bytes = len(stegoMessage_toBase64_bytes)
        stegoMessage_toBase64_size_bits = 8 * stegoMessage_toBase64_size_bytes
        #print("Stego-message Base64 bytes:", stegoMessage_toBase64_size_bytes)
        #print("Stego-message Base64 bits:", stegoMessage_toBase64_size_bits)

        embedded = False
        while not embedded:
            # Check if the paragraph has enough runs to embed the message
            print("Checking if the cover object is valid for embedding...")
            is_valid = is_capacity_enough_for_message(document, stegoMessage_toBase64_size_bits, None)
            print("The cover object is valid:", is_valid)
            if not is_valid:
                print("Not enough capacity in the document to embed the message.")
                break

            # Embed stego-message in DOCX
            print("Embedding stego-message...")
            random_paragraph_index = choose_random_paragraph(document, 'TODO', is_capacity_enough_for_message, stegoMessage_toBase64_size_bits)
            if random_paragraph_index is None:
                print("No paragraphs available for embedding.")
                break

            payload = stegoMessage_toBase64_size_bytes
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
                                next_stego_index = embedding_in_run(run, stegoMessage_toBase64_text, stego_index, payload)
                                stego_index = next_stego_index
                            else:
                                break
                else:
                    break
            
            #print("Extracting stego-message...")
            if stego_message_text != stego_message_extraction(document):
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