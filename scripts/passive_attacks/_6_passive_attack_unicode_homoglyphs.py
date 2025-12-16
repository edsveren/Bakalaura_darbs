from docx.document import Document as DocumentObject
import scripts.passive_attacks.unified_passive_attack_file as unified_passive_attack_file
from scripts.stego_methods._6_stego_method_unicode_homoglyphs import stego_message_extraction

# Passive attack
def stego_message_extraction_check(document: DocumentObject) -> str:
    try:
        return stego_message_extraction(document)
    except Exception as e:
        # print(e)
        return "TOO CORRUPT"

# Main    
def main(desired_file: str|None) -> None:
    unified_passive_attack_file.passive_attack("stego_method_6", stego_message_extraction_check, None, desired_file)
            
if __name__ == "__main__":
    main(None)
