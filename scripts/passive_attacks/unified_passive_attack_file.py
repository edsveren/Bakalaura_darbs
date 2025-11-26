import os
import csv
import difflib
from pathlib import Path
from collections import Counter
from typing import Callable
from itertools import zip_longest
from docx import Document
from docx.document import Document as DocumentObject

# Extract the text from the document
def extract_text(document: DocumentObject) -> str:
    text = []
    for paragraph in document.paragraphs:
        text.append(paragraph.text.replace('\xa0', '\x20')) # NBSP -> space
    # text = "\n".join(text)
    text = ''.join(text)
    #print(f"Text from document:'\n' {text}")
    return text

# Stego-message
def stego_message() -> str:
    stegoMessageText = Path("stego_messages/stego_message.txt").read_text(encoding="utf-8")
    #print(f"Stego-message: {stegoMessageText}")
    return stegoMessageText

# Perform an individual document check
def check_for_stego_message(
        stego_message_text: str, 
        stego_message_extracted: str,
        desired_file: str|None
    ) -> tuple[str, float]:

    # Make non-printable characters visible for analysis
    stego_message_extracted_readable = ''
    # Line feeds and carriage returns break the output format, so replace them with '?'
    for char in stego_message_extracted:
        if char.isprintable():
            stego_message_extracted_readable += char
        else:
            stego_message_extracted_readable += '�'

    # Print the original stego-message if desired file is specified
    if desired_file != None:
        print(f"Stego-message: [ {stego_message_text} ]")
    
    # Calculate the difference percentage between the original and extracted stego-message
    # Using difflib library which uses Ratcliff/Obershelp algorithm
    stego_message_difference = difflib.SequenceMatcher(None, stego_message_text, stego_message_extracted_readable)
    # Make it a percentage
    stego_message_difference_percentage = round(stego_message_difference.ratio() * 100, 2)

    # Determine the state of the stego-message based on the difference percentage
    if stego_message_difference_percentage != 0.0:
        # Print the extracted stego-message if desired file is specified
        if desired_file != None:
            print(f"Extracted message: [ {stego_message_extracted_readable} ]!")
            
        # STAGE 1: STEGO-MESSAGE GOOD
        if stego_message_difference_percentage == 100.0:
            state = "SAFE"
        else:
            # STAGE 2: STEGO-MESSAGE SLIGHTLY DEGRADED
            if stego_message_difference_percentage >= 95.0:
                state = "ALMOST SAFE (Less than 5% corruption)"
            # STAGE 3: STEGO-MESSAGE DEGRADED
            elif stego_message_difference_percentage >= 50.0:
                state = "SIGNIFICANTLY CORRUPTED (Up to 50% corruption)"
            # STAGE 4: STEGO-MESSAGE TOO DEGRADED 
            else:
                state = "HEAVILY CORRUPTED (Over 50% corruption)"
    else:
        # STAGE 5: THERE IS NO STEGO-MESSAGE
        if desired_file != None:
            print("The stego-message is gone!")
        state = "MISSING"
    
    # Print the difference percentage if desired file is specified
    if desired_file != None:
        print(f"Extracted message resemblance to the original stego-message: {stego_message_difference_percentage}%")

    return state, stego_message_difference_percentage

# Export to CSV for data analysis
def export_to_csv(
        stego_method_or_file: str,
        data_to_csv: list[list]
    ) -> None:

    # The non-transposed data result file
    result_file = f"results/passive_attacks/not_transposed/{stego_method_or_file}.csv"
    # The temporary file to store individual attack results before creating a regular file
    temporary_file = f"results/passive_attacks/not_transposed/temporary.csv"

    # Create temporary file to store individual attack results
    with open(temporary_file, "w", encoding="utf-8", newline="") as output_file:
        writer = csv.writer(output_file, delimiter=";")
        for row in data_to_csv:
            writer.writerow(row)
        writer.writerow('')

    # Transpose temporary file data in a more readable format
    export_to_csv_transposed(stego_method_or_file, temporary_file)

    if Path(result_file).is_file():
        # Read the temporary file data
        # Append the next attack results to the existing non-transposed result file
        with open(temporary_file, "r", encoding="utf-8", newline="") as input_file, open(result_file, "a+", encoding="utf-8", newline="") as output_file:
            reader = csv.reader(input_file, delimiter=";")
            writer = csv.writer(output_file, delimiter=";")
            for row in reader:
                writer.writerow(row)
    else:
        # Read the temporary file data
        # Create the non-transposed result file with the first attack type
        with open(temporary_file, "r", encoding="utf-8", newline="") as input_file, open(result_file, "w", encoding="utf-8", newline="") as output_file:
            reader = csv.reader(input_file, delimiter=";")
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Stego-method", stego_method_or_file])
            for row in reader:
                writer.writerow(row)
    
    # Delete the temporary file
    delete_file(Path(temporary_file))

    print(f"Created a CSV file: {str(Path(result_file))}")

# Transpose the temporary file data
def export_to_csv_transposed(stego_method_or_file: str, temporary_file: str):
    transposed_file = f"results/passive_attacks/transposed/{stego_method_or_file}_transposed.csv"
    
    # Read the temporary file data
    with open(temporary_file, "r", encoding="utf-8", newline="") as input_file:
        reader = csv.reader(input_file, delimiter=";")
        rows = list(reader)

    # Transpose using zip_longest to handle unequal row and columns lengths
    # Replace "missing cells" with empty strings
    transposed_rows = list(zip_longest(*rows, fillvalue=""))

    if Path(transposed_file).is_file():
        # Append the next attack results to the existing transposed result file
        with open(transposed_file, "a+", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            for row in transposed_rows:
                writer.writerow(row)
            writer.writerow('')            
    else:
        # Create the transposed result file with the first attack type
        with open(transposed_file, "w", encoding="utf-8", newline="") as output_file:
            writer = csv.writer(output_file, delimiter=";")
            writer.writerow(["Stego-method", stego_method_or_file])
            for row in transposed_rows:
                writer.writerow(row)
            writer.writerow('')

    print(f"Created a CSV file: {str(Path(transposed_file))}")

# Export all individual passive attack results to a single CSV file
def export_to_csv_all(desired_file: str|None) -> None:

    result_name = ''
    file_name = ''
    if desired_file != None:
        result_name = f'_{desired_file}'
        file_name = f'{desired_file}_'

    # The result file containing all individual passive attack results
    result_file = f"results/passive_attacks/transposed/unified_passive_attack_file{result_name}.csv"

    # Delete existing result file if exists
    delete_file(Path(result_file))

    # Loop through each individual passive attack result file and merge them
    i = 1
    for i in range(1, 7):

        # Access the individual passive attack transposed result file
        individual_transposed_file = f"results/passive_attacks/transposed/{file_name}stego_method_{i}_transposed.csv"
        if Path(individual_transposed_file).is_file():

            # Read the individual passive attack transposed result file data
            with open(individual_transposed_file, "r", encoding="utf-8", newline="") as input_file:
                reader = csv.reader(input_file, delimiter=";")
                individual_stego_rows = list(reader)
            
            # Read existing unified passive attack result file data if it exists
            if Path(result_file).is_file():
                with open(result_file, "r", encoding="utf-8", newline="") as input_file:
                    reader = csv.reader(input_file, delimiter=";")
                    existing_individual_stego_rows = list(reader)
            # Otherwise, create an empty list
            else:
                existing_individual_stego_rows = []

            # Append individual passive attack results to the unified passive attack result file
            with open(result_file, "w", encoding="utf-8", newline="") as output_file:
                writer = csv.writer(output_file, delimiter=";")
                # Create an index to track rows
                j = 0
                # Loop through each row in the individual passive attack result file
                for row in individual_stego_rows:
                    # If existing unified passive attack result file has data, merge them
                    if existing_individual_stego_rows:
                        # For the first row, add spacer columns for better readability
                        if j == 0:
                            number_of_columns = 4
                            spacer = [''] * number_of_columns
                            writer.writerow(existing_individual_stego_rows[j] + spacer + row[1:])
                        else:
                            writer.writerow(existing_individual_stego_rows[j] + row[1:])
                    # Otherwise, just write the individual passive attack result file data        
                    else:
                        writer.writerow(row)
                    j += 1

    print(f"Created a CSV file: {str(Path(result_file))}")

# Delete file
def delete_file(file: Path) -> None:
    if file.is_file() and not file.name.startswith("."):
        os.remove(file)
        print(f"Deleted: {str(file)}")

# Clean individual passive steganalysis attack results or specific file results
def clean_results_individual(attack_type_or_file: str) -> None:
    file = Path(f"results/passive_attacks/not_transposed/{attack_type_or_file}.csv")
    file_transposed = Path(f"results/passive_attacks/transposed/{attack_type_or_file}_transposed.csv")
    delete_file(file)
    delete_file(file_transposed)
    
### Main function ###
def passive_attack(
        stego_method: str, 
        stego_message_extraction: Callable[[DocumentObject], str],
        stego_message_text: str|None,
        desired_file: str|None
    ) -> None:
    
    # Clean up CSV files
    if desired_file == None:
        desired_file_csv_name = ''
        clean_results_individual(stego_method)
    else:
        desired_file_csv_name = f"{desired_file.rsplit('.', 1)[0]}_{stego_method}"
        clean_results_individual(desired_file_csv_name)

    # Attacked DOCX file data set
    attacked_stego_files = "data_set/attacked_stego_files"

    # Stego-message
    # Can sometimes be provided immediately for specific stego-methods
    if stego_message_text == None:
        stego_message_text = stego_message()

    print()
    # Loop through 10 active steganalysis directories
    for attack_directories in Path(attacked_stego_files).iterdir():
        
        # Current steganalysis attack directory
        attack_directory_name = attack_directories.name
        print(attack_directory_name)

        # The state of the stego-message and the number of DOCX files in the data
        stego_message_states_list = []
        stego_message_extracted_list = []
        stego_message_corruption_list = []
        nr_of_files = 0

        print(f"Stego-method: {stego_method}")
        print("Extracting stego-messages...")

        # Loop through each stego-method data set in attack directories
        for stego_directories in attack_directories.iterdir():

            # Choose only the attacked data sets for the specific stego-method
            if stego_directories.name == stego_method:
                desired_file_found = False

                # Loop through each individual file in the select stego-method data set
                for file in stego_directories.iterdir():

                    # Ignore temporary and git files
                    if file.is_file() and not file.name.startswith("."):

                        # Process either all files in the data set or only the desired file if specified
                        if desired_file == None or file.name == desired_file:
                            # Process only the desired file if specified
                            if desired_file != None and file.name == desired_file:
                                desired_file_found = True
                            
                            docPath = str(Path(f"{attacked_stego_files}/{attack_directories.name}/{stego_directories.name}/{file.name}"))
                            document = Document(docPath)
                            # Extract the stego-message from the document
                            # Using the provided extraction function for the specific stego-method
                            stego_message_extracted = stego_message_extraction(document)
                            # Analyze the stego-message in the individual document
                            stego_message_state, stego_message_difference_percentage = check_for_stego_message(stego_message_text, stego_message_extracted, desired_file)
                            # Store results for CSV export
                            stego_message_extracted_list.append(stego_message_extracted)
                            stego_message_corruption_list.append(stego_message_difference_percentage)
                            stego_message_states_list.append(stego_message_state)
                            nr_of_files += 1

                            if desired_file_found:
                                break

        if desired_file == None:
            # Count frequencies of each of three states
            counter = {key: Counter(stego_message_states_list).get(key, 0) for key in ['SAFE', 'ALMOST SAFE (Less than 5% corruption)', 'SIGNIFICANTLY CORRUPTED (Up to 50% corruption)', 'HEAVILY CORRUPTED (Over 50% corruption)', 'MISSING']}
            
            stego_message_states_list = list(counter.keys())
            state_frequencies = list(counter.values())
            frequency_percentages = []
            # Print results to terminal for the viewer
            for size, frequency in counter.items():
                frequency_percent = str(round((frequency / nr_of_files) * 100, 2)) #.replace(".", ",")
                print(f"State: {size}! Amount: {frequency} out of {nr_of_files} ({frequency_percent}%).")
                frequency_percentages.append(frequency_percent)

            # Export to CSV for data analysis
            data_to_csv = [
                ["Attack type", attack_directory_name],
                ["Number of documents attacked", nr_of_files],
                ["Stego-message state", *stego_message_states_list],
                ["Number of such states", *state_frequencies],
                ["Number of such states (%)", *frequency_percentages],
            ]
            export_to_csv(stego_method, data_to_csv)
        else:
            # Export to CSV for data analysis
            data_to_csv = [
                ["Attack type", attack_directory_name],
                ["Stego-message", stego_message_text],
                ["Extracted stego-message", *stego_message_extracted_list],
                ["Stego-message state", *stego_message_states_list],
                ["Stego-message integrity level", *stego_message_corruption_list]
            ]
            export_to_csv(desired_file_csv_name, data_to_csv)

        print()
        
    print("Extraction over!")
    print()

if __name__ == "__main__":
    export_to_csv_all(None)
