from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.oxml.parser import OxmlElement
import scripts.active_attacks.unified_active_attack_file as unified_active_attack_file

def retype_attack(
        stegoDocPath: str, 
        attackedStegoDocPath: str
        ) -> None:
    
    # Open the stego-file
    document = Document(stegoDocPath)

    # Create a new document
    new_document = Document()

    # Copy stego-document content and add it to the new document content
    print(f"Copying all content of: {Path(stegoDocPath).name} into a new document")
    # Loop through each paragraph
    for paragraph in document.paragraphs:
        # Create an empty paragraph element
        new_paragraph_element = OxmlElement('w:p')

        # Loop through each run
        for run in paragraph.runs:

            # Access the stego-file run element
            current_run_element = run._r
            # Create an empty run element
            new_run_element = OxmlElement('w:r')

            # Copy each stego-file run and its content into the new run
            for child_element in current_run_element:
                new_run_element.append(deepcopy(child_element))

            # Add the new run element to the new paragraph element
            new_paragraph_element.append(new_run_element)
        
        # Add the new paragraph element to the new document content
        new_document._element.body.append(new_paragraph_element)

    # Save
    new_document.save(attackedStegoDocPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("08_retype_attack", retype_attack, False)