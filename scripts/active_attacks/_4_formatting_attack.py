from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
import scripts.active_attacks.unified_active_attack_file as unified_active_attack_file

# Change typeface, font size and color
def formatting_attack(
        stegoDocPath: str, 
        attackedStegoDocPath: str
        ) -> None:
    
    document = Document(stegoDocPath)

    print(f"Changing the text font format: {Path(stegoDocPath).name}")
    for paragraph in document.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Britannic Bold'
            run.font.size = Pt(16)
            run.font.color.rgb = RGBColor(0x23, 0x23, 0x23)

    document.save(attackedStegoDocPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("04_format_attack", formatting_attack, False)