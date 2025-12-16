from docx.document import Document as DocumentObject
from scripts.stego_methods import unified_stego_file
import scripts.passive_attacks.unified_passive_attack_file as unified_passive_attack_file
from scripts.stego_methods._5_stego_method_unispace import stego_message_extraction, stego_message_standarization_to_unispace_method

# Passive attack
def stego_message_extraction_check(document: DocumentObject) -> str:
    try:
        return stego_message_extraction(document)
    except Exception as e:
        # print(e)
        return "TOO CORRUPT"

# Main
def main(desired_file: str|None) -> None:
    stego_message_text, _ = unified_stego_file.stego_message()
    stegoMessageInWhiteSpaceUnicode = stego_message_standarization_to_unispace_method(stego_message_text)
    unified_passive_attack_file.passive_attack("stego_method_5", stego_message_extraction_check, stegoMessageInWhiteSpaceUnicode, desired_file)
            
if __name__ == "__main__":
    main(None)