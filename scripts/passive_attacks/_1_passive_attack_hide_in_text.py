import base64
import unified_passive_attack_file
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    # Initialize an empty string to hold the stego-message in Base64
    stego_message_as_base64 = ''
    # Loop through all document paragraphs
    for paragraph in document.paragraphs:
        # Loop through all runs in the paragraph
        for run in paragraph.runs:
            # Get the run element's properties
            run_properties = run._r.rPr
            # Only analyze runs with run properties
            if run_properties != None:
                # Only analyze runs with exactly one character
                if len(run.text) == 1:
                    # Find the run property elements for color, size and hidden function
                    color_element = run_properties.find(qn('w:color'))
                    font_size_element = run_properties.find(qn('w:sz'))
                    vanish_element = run_properties.find(qn('w:vanish'))

                    # Check if all three run property elements are present
                    if None not in (color_element, font_size_element, vanish_element):
                        # Get their values
                        color_element_value = color_element.get(qn('w:val'))
                        font_size_value = font_size_element.get(qn('w:val'))
                        # And check if said values match the stego-embedding criteria
                        # Color = white (FFFFFF), size = 2 (1px), hidden = true (exists)
                        if color_element_value.upper() == 'FFFFFF' and font_size_value == '2':
                            # The run contains a stego-character
                            # Append it to the stego-message in Base64
                            stego_message_as_base64 += run.text
    try:
        #print(stegoMessage_as_base64)
        # Decode the stego-message from Base64 into UTF-8 (readable text)
        stego_message = base64.b64decode(stego_message_as_base64).decode('utf-8')
    except Exception as e: 
        #print(e)
        stego_message = "TOO CORRUPT"
    #print(stegoMessage)
    return stego_message

def main(desired_file: str|None) -> None:
    unified_passive_attack_file.passive_attack("stego_method_1", stego_message_extraction, None, desired_file)
            
if __name__ == "__main__":
    main(None)