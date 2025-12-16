from pathlib import Path
from win32com.client.dynamic import CDispatch as dynamic_CDispatch
import scripts.active_attacks.unified_active_attack_file as unified_active_attack_file

def save_as_attack(
        word: dynamic_CDispatch, 
        stegoDocPath: str, 
        attackedStegoDocPath: str, 
        wdFormatDocumentDefault: int
        ) -> None:
    
    # Simply Save the stego-file in a new location and Close
    print(f"Saving: {Path(stegoDocPath).name} as a new document")
    document = word.Documents.Open(stegoDocPath, ReadOnly=1, AddToRecentFiles=False)
    document.SaveAs2(attackedStegoDocPath, FileFormat=wdFormatDocumentDefault)
    document.Close()

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("06_save_as_attack", save_as_attack, True)