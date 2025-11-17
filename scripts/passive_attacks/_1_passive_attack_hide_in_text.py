import base64
import unified_passive_attack_file
from docx.oxml.ns import qn
from docx.document import Document as DocumentObject

# Extraction algorithm
def stego_message_extraction(document: DocumentObject) -> str:
    stegoMessage = ''
    stegoMessage_as_base64 = ''
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run_properties = run._r.rPr
            if run_properties != None:
                if len(run.text) == 1:
                    #base_64_char = re.search(r'[A-Za-z0-9+/]', run.text, flags=re.UNICODE)
                    color_element = run_properties.find(qn('w:color'))
                    font_size_element = run_properties.find(qn('w:sz'))
                    vanish_element = run_properties.find(qn('w:vanish'))

                    if None not in (color_element, font_size_element, vanish_element):
                        color_element_value = color_element.get(qn('w:val'))
                        font_size_value = font_size_element.get(qn('w:val'))
                        #vanish_element_value = vanish_element.get(qn('w:val'))
                        if color_element_value.upper() == 'FFFFFF' and font_size_value == '2': #and base_64_char != None:
                            stegoMessage_as_base64 += run.text
    try:
        #print(stegoMessage_as_base64)
        stegoMessage = base64.b64decode(stegoMessage_as_base64).decode('utf-8', errors='replace')
    except Exception as e: 
        #print(e)
        stegoMessage = "HEAVILY CORRUPTED (Over 50% corruption)"
    #print(stegoMessage)
    return stegoMessage

def main(desired_file: str|None) -> None:
    unified_passive_attack_file.passive_attack("stego_method_1", stego_message_extraction, None, desired_file)
            
if __name__ == "__main__":
    main(None)