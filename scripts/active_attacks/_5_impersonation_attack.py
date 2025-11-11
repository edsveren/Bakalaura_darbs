import os
from pathlib import Path
from win32com.client.dynamic import CDispatch as dynamic_CDispatch
import unified_active_attack_file

def impersonation_attack(
        word: dynamic_CDispatch, 
        stegoDocPath: str, 
        stegoPDFPath: str, 
        attackedStegoDocPath: str,
        wdFormatPDF: int,
        wdFormatDocumentDefault: int
        ) -> None:
    
    # Convert DOCX to PDF
    print(f"Converting: {Path(stegoDocPath).name} to: {Path(stegoPDFPath).name}")
    document = word.Documents.Open(stegoDocPath, ReadOnly=1, AddToRecentFiles=False)
    # Save PDF and close
    document.SaveAs2(stegoPDFPath, FileFormat=wdFormatPDF)
    document.Close()

    # Convert PDF back to DOCX
    print(f"Converting: {Path(stegoPDFPath).name} to: {Path(attackedStegoDocPath).name}")
    document = word.Documents.Open(stegoPDFPath, ReadOnly=1, AddToRecentFiles=False)
    # Save and Close
    document.SaveAs2(attackedStegoDocPath, FileFormat=wdFormatDocumentDefault)
    document.Close()

    # Remove any leftover PDF files
    if Path(stegoPDFPath).exists():
        os.remove(stegoPDFPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("05_impersonation_attack", impersonation_attack, True)