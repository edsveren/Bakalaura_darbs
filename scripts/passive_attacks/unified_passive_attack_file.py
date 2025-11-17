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
def analyse_corruption_level(
        stego_message_text: str, 
        stego_message_extracted: str,
        desired_file: str|None,
        ) -> float:
    # Calculate the difference percentage between the original and extracted stego-message
    # Using difflib library which uses Ratcliff/Obershelp algorithm
    stego_message_difference = difflib.SequenceMatcher(None, stego_message_text, stego_message_extracted)
    # Make it a percentage
    stego_message_difference_percentage = round(stego_message_difference.ratio() * 100, 2)

    # With above 95% resemblance, the stego-message is still very readable
    # if stego_message_difference_percentage >= 95.0:
    #     return stego_message_difference_percentage
    
    if desired_file != None:
        if stego_message_extracted != "HEAVILY CORRUPTED (Over 50% corruption)":
            # The stego-message was corrupted but still extractable
            print("Stego-message should be:")
            print(stego_message_text)
            print("But instead is:")
            print(stego_message_extracted)
            # print(f"Corrupted message resemblance to the original stego-message: {stego_message_difference_percentage}%")
        else:
            # The stego-message was too corrupted to be extracted
            print("Stego-message too corrupted to be extracted properly!")
        print(f"Corrupted message resemblance to the original stego-message: {stego_message_difference_percentage}%")
    return stego_message_difference_percentage

# An individual document check
def check_for_stego_message(
        file_name: str, 
        document: DocumentObject, 
        stego_message_text: str, 
        # stego_message_extraction: Callable[[DocumentObject], str]
        stego_message_extracted: str,
        desired_file: str|None
    ) -> tuple[str, float|str]:

    # Extract the stego-message from the document
    # Using the provided extraction function for the specific stego-method
    # stego_message_extracted = stego_message_extraction(document)

    if stego_message_extracted != '':
        #print(stego_message_extracted)
        if stego_message_text == stego_message_extracted:
            #print(f"{file_name}'s extracted stego-message: {stego_message_extracted}. EQUAL!")
            return "SAFE", 100.0 # STAGE 1: STEGO-MESSAGE GOOD
        else:
            # Calculate the corruption level
            stego_message_difference_percentage = analyse_corruption_level(stego_message_text, stego_message_extracted, desired_file)
            if stego_message_difference_percentage >= 95.0:
                return "ALMOST SAFE (Less than 5% corruption)", stego_message_difference_percentage # STAGE 2: STEGO-MESSAGE SLIGHTLY DEGRADED
            elif stego_message_difference_percentage >= 50.0:
                return "SIGNIFICANTLY CORRUPTED (Up to 50% corruption)", stego_message_difference_percentage # STAGE 3: STEGO-MESSAGE DEGRADED
            else:
                return "HEAVILY CORRUPTED (Over 50% corruption)", stego_message_difference_percentage # STAGE 4: STEGO-MESSAGE TOO DEGRADED
    else:
        #print(f"{file_name}'s extracted stego-message: THERE IS NO STEGO-MESSAGE!")
        return "MISSING", 'NA' # STAGE 3: THERE IS NO STEGO-MESSAGE
    
# def export_passive_attack_to_csv() -> None:

# Export to CSV for data analysis
def export_to_csv(
        stego_method_or_file: str,
        # attack_type: str, 
        # number_of_docs: int, 
        # states: list, 
        # state_frequencies: list, 
        # state_frequency_percentages: list
        data_to_csv: list[list]
    ) -> None:

    # The non-transposed data result file
    result_file = f"results/passive_attacks/not_transposed/{stego_method_or_file}.csv"
    # The temporary file to store individual attack results before creating a regular file
    temporary_file = f"results/passive_attacks/not_transposed/temporary.csv"

    # Create temporary file to store individual attack results
    with open(temporary_file, "w", encoding="utf-8", newline="") as output_file:
        writer = csv.writer(output_file, delimiter=";")
        # writer.writerow(["Attack type", attack_type])
        # writer.writerow(["Number of documents attacked", number_of_docs])
        # writer.writerow(["Stego-message state", *states])
        # writer.writerow(["Number of such states", *state_frequencies])
        # writer.writerow(["Number of such states (%)", *state_frequency_percentages])
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
    with open(temporary_file, newline="", encoding="utf-8") as input_file:
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
def export_to_csv_all() -> None:

    # The result file containing all individual passive attack results
    result_file = f"results/passive_attacks/transposed/unified_passive_attack_file.csv"

    # Delete existing result file if exists
    delete_file(Path(result_file))

    # Loop through each individual passive attack result file and merge them
    i = 1
    for i in range(1, 7):

        # Access the individual passive attack transposed result file
        individual_transposed_file = f"results/passive_attacks/transposed/stego_method_{i}_transposed.csv"
        if Path(individual_transposed_file).is_file():

            # Read the individual passive attack transposed result file data
            with open(individual_transposed_file, "r", encoding="utf-8", newline="") as input_file:
                reader = csv.reader(input_file, delimiter=";")
                individual_stego_rows = list(reader)
            
            # Read existing unified passive attack result file data if exists
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
                j = 0
                for row in individual_stego_rows:
                    if existing_individual_stego_rows:
                        if j == 0:
                            number_of_columns = 4
                            spacer = [''] * number_of_columns
                            writer.writerow(existing_individual_stego_rows[j] + spacer + row[1:])
                        else:
                            #spacer = [""]
                            writer.writerow(existing_individual_stego_rows[j] + row[1:])
                    else:
                        writer.writerow(row)
                    j += 1
    print(f"Created a CSV file: {str(Path(result_file))}")

# Delete file
def delete_file(file: Path) -> None:
    if file.is_file() and not file.name.startswith("."):
        print(f"Deleting: {str(file)}")
        os.remove(file)

# Clean individual passive steganalysis attack results
def clean_results_individual(attack_type: str) -> None:
    file = Path(f"results/passive_attacks/not_transposed/{attack_type}.csv")
    file_transposed = Path(f"results/passive_attacks/transposed/{attack_type}_transposed.csv")
    delete_file(file)
    delete_file(file_transposed)

# Clean specific file from passive steganalysis attack results
def clean_results_specific_file(desired_file: str) -> None:
    print("TODO")
    
### Main function ###
def passive_attack(
        stego_method: str, 
        stego_message_extraction: Callable[[DocumentObject], str],
        stego_message_text: str|None,
        desired_file: str|None
    ) -> None:
    
    # Clean up CSV files
    if desired_file == None:
        clean_results_individual(stego_method)
    else:
        clean_results_specific_file(desired_file)

    # Attacked DOCX file data set
    attacked_stego_files = "data_set/attacked_stego_files"

    # Stego-message
    if stego_message_text == None:
        stego_message_text = stego_message()

    stego_message_extracted = ''

    print()
    # Loop through 10 active steganalysis directories
    for attack_directories in Path(attacked_stego_files).iterdir():
        
        # Current steganalysis attack directory
        attack_directory_name = attack_directories.name
        print(attack_directory_name)

        # The state of the stego-message and the number of DOCX files in the data
        states_list = []
        corruption_list = []
        nr_of_files = 0

        print("Extracting stego-messages...")

        # Loop through each stego-method data set in attack directories
        for stego_directories in attack_directories.iterdir():

            # Choose only the attacked data sets for the specific stego-method
            if stego_directories.name == stego_method:

                # Loop through each individual file in the select stego-method data set
                for file in stego_directories.iterdir():
                    desired_file_found = False

                    # Ignore temporary and git files
                    if file.is_file() and not file.name.startswith("."):
                        # Process only the desired file if specified
                        if desired_file != None and file.name == desired_file:
                            docPath = str(Path(f"{attacked_stego_files}/{attack_directories.name}/{stego_directories.name}/{desired_file}"))
                            desired_file_found = True
                        # Otherwise, process all files in the data set
                        else:
                            docPath = str(Path(f"{attacked_stego_files}/{attack_directories.name}/{stego_directories.name}/{file.name}"))

                        # Return the state of the stego-message in the individual document
                        document = Document(docPath)
                        # Extract the stego-message from the document
                        # Using the provided extraction function for the specific stego-method
                        stego_message_extracted = stego_message_extraction(document)

                        state, stego_message_difference_percentage = check_for_stego_message(file.name, document, stego_message_text, stego_message_extracted, desired_file)
                        corruption_list.append(stego_message_difference_percentage)
                        states_list.append(state)
                        nr_of_files += 1

                    if desired_file_found:
                        break    

            # if desired_file_found:
            #     break

        # Count frequencies of each of three states
        counter = {key: Counter(states_list).get(key, 0) for key in ['SAFE', 'ALMOST SAFE (Less than 5% corruption)', 'SIGNIFICANTLY CORRUPTED (Up to 50% corruption)', 'HEAVILY CORRUPTED (Over 50% corruption)', 'MISSING']}
        
        states_list = list(counter.keys())
        state_frequencies = list(counter.values())
        frequency_percentages = []
        
        if desired_file == None:
            # Print results to terminal for the viewer
            for size, frequency in counter.items():
                frequency_percent = str(round((frequency / nr_of_files) * 100, 2)) #.replace(".", ",")
                print(f"State: {size}! Amount: {frequency} out of {nr_of_files} ({frequency_percent}%).")
                frequency_percentages.append(frequency_percent)

            data_to_csv = [
                ["Attack type", attack_directory_name],
                ["Number of documents attacked", nr_of_files],
                ["Stego-message state", *states_list],
                ["Number of such states", *state_frequencies],
                ["Number of such states (%)", *frequency_percentages],
            ]
            export_to_csv(stego_method, data_to_csv)
        else:
            data_to_csv = [
                ["Attack type", attack_directory_name],
                ["Stego-message", stego_message_text],
                ["Extracted stego-message", stego_message_extracted],
                ["Stego-message state", *states_list],
                ["Stego-message integrity level", *corruption_list]
            ]
            # export_to_csv(desired_file, data_to_csv)

        # data_to_csv = {
        #     "Attack type": attack_directory_name,
        #     "Number of documents attacked": nr_of_files,
        #     "Stego-message state": states_list,
        #     "Number of such states": state_frequencies,
        #     "Number of such states (%)": frequency_percentages
        # }

        # Export to CSV for data analysis
        # export_to_csv(stego_method, attack_directory_name, nr_of_files, states_list, state_frequencies, frequency_percentages)
        print()
        
    print("Extraction over!")

if __name__ == "__main__":
    export_to_csv_all()
