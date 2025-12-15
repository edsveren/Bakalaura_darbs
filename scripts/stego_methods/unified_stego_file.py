import random
from pathlib import Path
from typing import Callable, Any
from docx import Document
from docx.text.run import Run
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn

# Get stego-message text and bytes from file
def stego_message() -> tuple[str, bytes]:
    stego_message_text = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    stego_message_bytes = stego_message_text.encode("utf-8")
    return stego_message_text, stego_message_bytes

# Extract text from the document
def extract_text(
        document: DocumentObject,
        NBSP: bool, # Non-Breaking Space
        index: int
    ) -> str:
    text = []
    # Loop through paragraphs starting from index
    for paragraph in document.paragraphs[index:]:
        # Append paragraph text, replace NBSP with space if specified
        if NBSP == True:
            text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
        else:
            text.append(paragraph.text)
    # Join all paragraph texts into a single string
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

    # Get all paragraphs in the document
    paragraphs = document.paragraphs
    # Return None if there are no paragraphs
    if not paragraphs:
        return None
    
    # Keep looking for a valid paragraph
    # Until a paragraph with enough capacity is found
    while True:
        # Select a random paragraph index within the paragraphs range
        # Using random number generator
        random_paragraph_index = random.randint(0, len(paragraphs) - 1)
        # Pick the random paragraph
        random_paragraph = paragraphs[random_paragraph_index]
        # Count necessary elements (words, chars, black chars, etc.) from the random paragraph index
        necessary_element_count = necessary_element_count_function(document, random_paragraph_index)
        # Check if the chosen paragraph has enough capacity for the stego-message
        is_valid = is_capacity_enough_for_message(document, necessary_element_count, stego_message_size_bits, random_paragraph_index)
        # If capacity is enough, return the paragraph index, ending the search
        if is_valid:
            print(f"Random paragraph start: {random_paragraph.text}")
            return random_paragraph_index

# Get stego-message bytes and bits     
def get_bytes_and_bits(stego_message_bytes: bytes|str) -> tuple[int, int]:
    # Get stego-message size in bytes 
    # By calculating the length of the corresponding string or bytes object
    stego_message_size_bytes = len(stego_message_bytes)
    # Get stego-message size in bits
    stego_message_size_bits = 8 * stego_message_size_bytes
    
    # print(f"Stego-message bytes: {stego_message_size_bytes}")
    # print(f"Stego-message bites: {stego_message_size_bits}")

    return stego_message_size_bytes, stego_message_size_bits

### Main file ###   
def stego_method(
        stego_method: str, # 'stego_method_1', 'stego_method_2', etc.
        specialized_stego_message: tuple[str, bytes]|None, # For specialized stego-message input
        necessary_element_count_function: Callable[[DocumentObject, int], int], # Count words, chars, black chars, etc.
        is_capacity_enough_for_message:  Callable[[DocumentObject, int, int, int], bool], # Check document capacity function
        embedding_algorithm: Callable[[Run, Any, int, int], int], # Embedding algorithm function, Any = str|bytes
        embedding_data_type: str, # 'string' or 'bytes'
        extraction_algorithm: Callable[[DocumentObject], str] # Extract stego-message function
    ) -> None:
    # Clean DOCX file directory
    base = "data_set/clean_files"
    # Loop through all DOCX files in the clean directory
    for file in Path(base).iterdir():
        docPath = str(Path(f"{base}/{file.name}")) #docPath=Path("data_set/clean_files/TEST_0.docx")
        print(f"DOCX file: {docPath}")
        print("Beginning the embedding process...")
        # Load the DOCX file document object
        document = Document(docPath)
        # Extract the full document text
        text = extract_text(document, True, 0)
        # Count the full corresponding document elements (words, chars, black chars, etc.)
        necessary_element_count = necessary_element_count_function(document, 0)

        # Get stego-message text and bytes from default file or specialized input
        if specialized_stego_message == None:
            stego_message_text, stego_message_bytes = stego_message()
        else:
            stego_message_text, stego_message_bytes = specialized_stego_message
        
        # Get stego-message size in bytes and bits
        # Unispace stego-method adds padding to the stego-message
        # So it requires more data to embed
        if stego_method == 'stego_method_5':
            stego_message_size_bytes, stego_message_size_bits = get_bytes_and_bits(stego_message_text)
        else:
            stego_message_size_bytes, stego_message_size_bits = get_bytes_and_bits(stego_message_bytes)

        # Embed stego-message until it is embedded successfully
        embedded = False
        while not embedded:
            # Check if the entire document has enough runs to embed the message
            print("Checking if the cover object is valid for embedding...")
            is_valid = is_capacity_enough_for_message(document, necessary_element_count, stego_message_size_bits, 0)
            print("The cover object is valid:", is_valid)

            # If it doesn't have enough capacity, break the loop
            if not is_valid:
                print("Not enough capacity in the document to embed the message.")
                break

            # Embed stego-message in DOCX
            print("Embedding stego-message...")
            # Pick a random paragraph index to start embedding
            random_paragraph_index = choose_random_paragraph(document, necessary_element_count_function, is_capacity_enough_for_message, stego_message_size_bits)
            # This should not happen since the capacity was already determined to be enough but just in case
            if random_paragraph_index == None:
                print("No paragraphs available for embedding.")
                break
            
            # Initialize the starting index and payload for embedding
            # Unicode homoglyphs stego-method adds a '1' at the start of the stego-message bits
            # To allow for proper extraction, so the payload needs to be increased by 1 bit
            if stego_method == 'stego_method_6':
                payload = stego_message_size_bits + 1
            else:
                payload = stego_message_size_bytes
            stego_index = 0

            # Determine the data to embed based on the specified type (most use string)
            data_to_embed = ''
            if embedding_data_type == 'bytes':
                data_to_embed = stego_message_bytes
            elif embedding_data_type == 'string':
                data_to_embed = stego_message_text

            # Loop through paragraphs starting from the chosen random paragraph index
            for paragraph in document.paragraphs [random_paragraph_index:]:
                # Check if the entire stego-message has been embedded
                # By comparing the current stego index with the payload size
                if stego_index < payload:
                    # Get the original amount of runs in the paragraph
                    # Because runs will be added or otherwise modified during embedding
                    # Which leads to endless loops or skipped runs
                    original_run_amount = list(paragraph.runs)
                    # Loop through each run in the original amount of runs
                    for run in original_run_amount:
                        # Get the run element to check if it contains text
                        run_element = run._r
                        # Only process runs that contain text elements
                        if run_element.find(qn('w:t')) != None:
                            # Check again if the entire stego-message has already been embedded at the current run
                            if stego_index < payload:
                                # The main embedding function call
                                # It works within each individual run
                                # To embed as much of the stego-message as possible in that run
                                # And returns the updated stego index after embedding 
                                stego_index = embedding_algorithm(run, data_to_embed, stego_index, payload)
                            # Break the inner loop if the entire stego-message has been embedded
                            else:
                                break
                # Break the outer loop if the entire stego-message has been embedded
                else:
                    break
            
            # Get the original stego-message for comparison
            # Unispace method is different
            # It compromises the original stego-message for embedding purposes
            # First off, it can only use uppercase ASCII characters
            # So all lowercase ASCII characters are converted to uppercase 
            # While non-ASCII characters are removed
            # And secondly, it uses whitespace padding characters at the start and end
            # To denote the start and end of the stego-message for the purposes of extraction
            if stego_method == 'stego_method_5':
                stego_message_text = stego_message_text[1:-1]
            else:
                stego_message_text, _ = stego_message()

            # Extract the stego-message from the document
            # print("Extracting stego-message...")
            extracted_stego_message = extraction_algorithm(document)
            print(f"Stego-message: {stego_message_text}")
            print(f"Extracted stego-message: {extracted_stego_message}")
            
            # Compare the extracted stego-message with the one before embedding
            # If they don't match, something went wrong with the embedding process
            # This breaks the loop and moves on to the next file
            # However, this is not an intended behavior
            # And the corresponding stego-method script must be fixed
            if stego_message_text != extracted_stego_message:
                print("Extracted message is not equal to stego-message!")
                break
            #print("Extraction successful!")
            print("Embedding successful!")
            # If everything went well, set embedded to True to exit the loop
            embedded = True

        # If everything was embedded successfully, save the modified document
        if embedded:
            stegoDocPath = str(Path(f"data_set/stego_files/{stego_method}/{file.name}"))
            document.save(stegoDocPath)
            print(f"Saved: {stegoDocPath}")
        # If, for some reason, embedding was not successful
        # Let the user know
        else:
            print("Embedding not possible.")
        print()