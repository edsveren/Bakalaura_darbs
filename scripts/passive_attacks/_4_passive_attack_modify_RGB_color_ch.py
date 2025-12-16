import unified_passive_attack_file
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    # Initialize an empty bytes object to hold the stego-message bytes
    stego_message_bytes = b''
    # Loop through all document paragraphs
    for paragraph in document.paragraphs:
        # Loop through all runs in the paragraph
        for run in paragraph.runs:
            # Get the run element's properties
            run_properties = run._r.rPr
            # Only analyze runs with run properties
            if run_properties != None:
                # Find the color element in the run properties
                color_element = run_properties.find(qn('w:color'))
                # And check if it is present
                if color_element != None:
                    # Get its value
                    color_element_value = color_element.get(qn('w:val'))
                    # And check if it is not black or automatic (default)
                    # In theory, any coloured text can be used
                    # But in practice, it requires some kind of stego-key
                    # Because it's difficult to extract the stego-message deterministically
                    # Without said key which provides original base values of non-stego-data
                    if color_element_value.lower() not in ('000000', 'auto'):
                        # Get bits from each RGB color channel's hex value
                        r_bit = int(color_element_value[0:2], 16)
                        g_bit = int(color_element_value[2:4], 16)
                        b_bit = int(color_element_value[4:6], 16)
                        # Red and Green color channels can contain max 111 (binary) = 7 (decimal) value
                        # Blue color channel can contain max 11 (binary) = 3 (decimal) value
                        if r_bit <= 7 and g_bit <= 7 and b_bit <= 3:
                            # Combine the bits into a single 8-bit string
                            stego_bit_string = f"{r_bit:03b}{g_bit:03b}{b_bit:02b}"
                            # And transform it into a byte
                            stego_byte = int(stego_bit_string, 2).to_bytes(1, 'big')
                            # print(stego_byte)
                            # Add the stego-byte to the stego-message bytes object
                            stego_message_bytes += stego_byte
    # print(stego_message_bytes)
    # Decode the stego-message from raw bytes into UTF-8 (readable text)
    stego_message = stego_message_bytes.decode('utf-8')
    # print(stego_message)
    return stego_message
            
def main(desired_file: str|None) -> None:
    unified_passive_attack_file.passive_attack("stego_method_4", stego_message_extraction, None, desired_file)
            
if __name__ == "__main__":
    main(None)
