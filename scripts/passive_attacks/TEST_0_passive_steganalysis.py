import os
import csv
import difflib
from pathlib import Path
from collections import Counter
from typing import Callable
from itertools import zip_longest
from docx import Document
from docx.document import Document as DocumentObject
import unified_passive_attack_file
import _1_passive_attack_hide_in_text
import _4_passive_attack_modify_RGB_color_ch
import _5_passive_attack_unispace
import _6_passive_attack_unicode_homoglyphs

def passive_attack(
        stego_method: str, 
        stego_message_extraction: Callable[[DocumentObject], str],
        stego_message_text: str|None,
        desired_file: str|None
    ) -> None:
    
    # Clean up CSV files
    unified_passive_attack_file.clean_results_individual(stego_method)

    # Attacked DOCX file data set
    attacked_stego_files = "data_set/attacked_stego_files"

    # Stego-message
    if stego_message_text == None:
        stego_message_text = unified_passive_attack_file.stego_message()

    print()
    # Loop through 10 active steganalysis directories
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

            # Choose only the attacked data sets for the specific stego-method
            if stego_directories.name == stego_method:

                # Loop through each individual file in the select stego-method data set
                for file in stego_directories.iterdir():
                    # Ignore temporary and git files
                    if file.is_file() and not file.name.startswith("."):
                        # Process only the desired file if specified
                        if desired_file != None and file.name == desired_file:
                            docPath = str(Path(f"{attacked_stego_files}/{attack_directories.name}/{stego_directories.name}/{file.name}"))
                            document = Document(docPath)
                            state = unified_passive_attack_file.check_for_stego_message(file.name, document, stego_message_text, stego_message_extraction)
                        # Otherwise, process all files in the data set
                        else:
                            # Return the state of the stego-message in the individual document
                            docPath = str(Path(f"{attacked_stego_files}/{attack_directories.name}/{stego_directories.name}/{file.name}"))
                            document = Document(docPath)
                            state = unified_passive_attack_file.check_for_stego_message(file.name, document, stego_message_text, stego_message_extraction)
                            states_list.append(state)
                            nr_of_files += 1

        # Count frequencies of each of three states
        counter = {key: Counter(states_list).get(key, 0) for key in ['SAFE', 'ALMOST SAFE (Less than 5% corruption)', 'CORRUPTED', 'MISSING']}
        
        if desired_file == None:
            states_list = list(counter.keys())
            state_frequencies = list(counter.values())
            frequency_percentages = []

            # Print results to terminal for the viewer
            for size, frequency in counter.items():
                frequency_percent = str(round((frequency / nr_of_files) * 100, 2)) #.replace(".", ",")
                print(f"State: {size}! Amount: {frequency} out of {nr_of_files} ({frequency_percent}%).")
                frequency_percentages.append(frequency_percent)
            
            # Export to CSV for data analysis
            unified_passive_attack_file.export_to_csv(stego_method, attack_directory_name, nr_of_files, states_list, state_frequencies, frequency_percentages)
            print()
        
    print("Extraction over!")

if __name__ == "__main__":

    desired_file = "TEST_0.docx"
    # Choose stego-method to attack
    stego_method = "stego_method_1"
    stego_message_extraction = _1_passive_attack_hide_in_text.stego_message_extraction
    stego_message_text = unified_passive_attack_file.stego_message()
    passive_attack(stego_method, stego_message_extraction, stego_message_text, desired_file)

    stego_method = "stego_method_4"
    stego_message_extraction = _4_passive_attack_modify_RGB_color_ch.stego_message_extraction
    stego_message_text = unified_passive_attack_file.stego_message()
    passive_attack(stego_method, stego_message_extraction, stego_message_text, desired_file)

    stego_method = "stego_method_5"
    stego_message_extraction = _5_passive_attack_unispace.stego_message_extraction
    stego_message_text = unified_passive_attack_file.stego_message()
    passive_attack(stego_method, stego_message_extraction, stego_message_text, desired_file)

    stego_method = "stego_method_6"
    stego_message_extraction = _6_passive_attack_unicode_homoglyphs.stego_message_extraction
    stego_message_text = unified_passive_attack_file.stego_message()
    passive_attack(stego_method, stego_message_extraction, stego_message_text, desired_file)