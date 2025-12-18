from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
import scripts.active_attacks.unified_active_attack_file as unified_active_attack_file

def formatting_attack(
        stegoDocPath: str, 
        attackedStegoDocPath: str
        ) -> None:
    
    # Open the stego-file
    document = Document(stegoDocPath)

    print(f"Changing the text font format: {Path(stegoDocPath).name}")
    # Loop through each paragraph
    for paragraph in document.paragraphs:
        # Loop through each run
        for run in paragraph.runs:
            # Change the run font theme, size and colour
            run.font.name = 'Britannic Bold'
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(0x23, 0x23, 0x23) # Black

    # Save
    document.save(attackedStegoDocPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("04_format_attack", formatting_attack, False)