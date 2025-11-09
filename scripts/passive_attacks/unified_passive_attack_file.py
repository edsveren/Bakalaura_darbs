import os
import csv
import pandas as pd
from pathlib import Path
from collections import Counter
from typing import Callable
from docx import Document
from docx.document import Document as DocumentObject

# Extract text from the document
def extract_text(document: DocumentObject) -> str:
    text = []
    for paragraph in document.paragraphs:
        text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
    text = "\n".join(text)
    text = ''.join(text)
    #print(f"Text from document:'\n' {text}")
    return text

# Stego-message
def stego_message() -> str:
    stegoMessageText = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    #print(f"Stego-message: {stegoMessageText}")
    return stegoMessageText

# Analyze each corrupt stego-message's corruption level
# def analyse_corruption_level(stego_message_text: str, stego_message_extracted: str):
#     i = 0
#     not_equal = 0
#     for char in stego_message_text:
#         if i > len(stego_message_extracted):
#             break
#         if char != stego_message_extracted[i]:
#             not_equal += 1
#         i += 1

# An individual document check
def check_for_stego_message(
        file_name: str, 
        document: DocumentObject, 
        stego_message_text: str, 
        stego_message_extraction: Callable[[DocumentObject], str]
    ) -> str:
    stego_message_extracted = stego_message_extraction(document)
    if stego_message_extracted != '':
        #print(stego_message_extracted)
        if stego_message_text == stego_message_extracted:
            #print(f"{file_name}'s extracted stego-message: {stego_message_extracted}. EQUAL!")
            return "SAFE" # STEGO-MESSAGE GOOD
        else:
            #analyse_corruption_level(stego_message_text, stego_message_extracted)
            #print(f"{file_name}'s extracted stego-message: {stego_message_extracted}. NOT EQUAL!")
            return "CORRUPTED" # STAGE 2: STEGO-MESSAGE DEGRADED
    else:
        #print(f"{file_name}'s extracted stego-message: THERE IS NO STEGO-MESSAGE!")
        return "MISSING" # THERE IS NO STEGO-MESSAGE

# Export to CSV for data analysis
def export_to_csv(
        stego_method: str, 
        attack_type: str, 
        number_of_docs: int, 
        states: list, 
        state_frequencies: list, 
        state_frequency_percentages: list
    ) -> None:
    result_file = f"results/passive_attacks/{stego_method}.csv"

    if Path(result_file).is_file():
        with open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Attack type", attack_type])
            writer.writerow(["Number of documents attacked", number_of_docs])
            writer.writerow(["Stego-message state", *states])
            writer.writerow(["Number of such states", *state_frequencies])
            writer.writerow(["Number of such states (%)", *state_frequency_percentages])
            writer.writerow('')
    else:
        with open(result_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Attack type", attack_type])
            writer.writerow(["Number of documents attacked", number_of_docs])
            writer.writerow(["Stego-message state", *states])
            writer.writerow(["Number of such states", *state_frequencies])
            writer.writerow(["Number of such states (%)", *state_frequency_percentages])
            writer.writerow('')

### TODO
### must create the same csv but from the perspective of the active attack, not stego-method
# def combine_all_csv_into_one():
#     data = {'Name': ['ANSH', 'VANSH'], 'Age': [25, 30]}
#     df = pd.DataFrame(data)
#     print("Original DataFrame:")
#     print(df)
#     transposed_df = df.transpose()
#     print("\nTransposed DataFrame:")
#     print(transposed_df)
#     print("NOT DONE")

### Main function ###
def passive_attack(
        stego_method: str, 
        stego_message_extraction: Callable[[DocumentObject], str],
        stego_message_text: str|None
    ) -> None:
    
    # CSV file
    csv_file = Path(f"results/passive_attacks/{stego_method}.csv")
    if csv_file.is_file():
        os.remove(csv_file)

    # Attacked DOCX file data set
    attacked_stego_files = "data_set/attacked_stego_files"

    # Stego-message
    if stego_message_text == None:
        stego_message_text = stego_message()

    print()
    # Loop through 10 active steganalysis attack directories
    for attack_directories in Path(attacked_stego_files).iterdir():
        
        # Current steganalysis attack directory
        attack_directory_name = attack_directories.name
        print(attack_directory_name)

        # The state of the stego-message and the number of DOCX files in the data
        states_list = []
        nr_of_files = 0

        print("Extracting stego-messages...")

        # Loop through each stego-method data set in attack directories
        for stego_directories in attack_directories.iterdir():
            #print(stego_directories.name)
            # Choosing only the attacked data sets for the specific stego-method
            if stego_directories.name == stego_method:

                # Loop through each individual file in the select stego-method data set
                for file in stego_directories.iterdir():
                    if file.is_file() and not file.name.startswith("."):
                        docPath = str(Path(f"{attacked_stego_files}/{attack_directories.name}/{stego_directories.name}/{file.name}"))
                        document = Document(docPath)
                        state = check_for_stego_message(file.name, document, stego_message_text, stego_message_extraction)
                        states_list.append(state)
                        nr_of_files += 1

        # Count frequencies of each of three states
        counter = {key: Counter(states_list).get(key, 0) for key in ['SAFE', 'CORRUPTED', 'MISSING']}

        states_list = list(counter.keys())
        state_frequencies = list(counter.values())
        frequency_percentages = []

        # Print results to terminal for the viewer
        for size, frequency in counter.items():
            frequency_percent = str(round((frequency / nr_of_files) * 100, 2)) #.replace(".", ",")
            print(f"State: {size}! Amount: {frequency} out of {nr_of_files} ({frequency_percent}%).")
            frequency_percentages.append(frequency_percent)
        
        # Export to CSV for data analysis
        export_to_csv(stego_method, attack_directory_name, nr_of_files, states_list, state_frequencies, frequency_percentages)
        print()
        
    print("Extraction over!")