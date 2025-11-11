from pathlib import Path
from win32com.client.dynamic import CDispatch as dynamic_CDispatch
import unified_active_attack_file

def document_inspect_attack(
        word: dynamic_CDispatch, 
        stegoDocPath: str, 
        attackedStegoDocPath: str, 
        wdFormatDocumentDefault: int,
        wdRDIAll: int
        ) -> None:
    
    print(f"Using Document Inspect tool on: {Path(stegoDocPath).name}")
    document = word.Documents.Open(stegoDocPath, ReadOnly=1, AddToRecentFiles=False)

    # Document Inspector Attack for Metadata
    document.RemoveDocumentInformation(RemoveDocInfoType=wdRDIAll)

    # Document Inspector Attack for Content
    # Each content needs a separate attack
    for i in range(1, document.DocumentInspectors.Count + 1):
        # DocumentInspector.Inspect method
        status, result = document.DocumentInspectors.Item(i).Inspect()
        print(f"{i}. {document.DocumentInspectors.Item(i).Name} Status: {status} Result: {result}")
        # DocumentInspector.Fix method
        if status == 1: 
            document.DocumentInspectors.Item(i).Fix()

    # Save and Close
    document.SaveAs2(attackedStegoDocPath, FileFormat=wdFormatDocumentDefault)
    document.Close()

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("10_document_inspector_attack", document_inspect_attack, True)