from pathlib import Path
from win32com.client.dynamic import CDispatch as dynamic_CDispatch
import scripts.active_attacks.unified_active_attack_file as unified_active_attack_file

def clear_format_attack(
        word: dynamic_CDispatch, 
        stegoDocPath: str, 
        attackedStegoDocPath: str, 
        wdFormatDocumentDefault: int
        ) -> None:
    
    print(f"Clearing all formatting of: {Path(stegoDocPath).name}")
    document = word.Documents.Open(stegoDocPath, ReadOnly=1, AddToRecentFiles=False)

    # Press CTRL + A
    word.Selection.WholeStory()

    # Click Clear All Formatting button
    word.Selection.ClearFormatting()

    document.SaveAs2(attackedStegoDocPath, FileFormat=wdFormatDocumentDefault)
    document.Close()

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("09_clear_format_attack", clear_format_attack, True)