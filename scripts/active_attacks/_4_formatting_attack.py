from pathlib import Path
from typing import cast
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor
from docx.styles.style import CharacterStyle
import scripts.active_attacks.unified_active_attack_file as unified_active_attack_file

def formatting_attack(
        stegoDocPath: str, 
        attackedStegoDocPath: str
        ) -> None:
    
    # Open the stego-file
    document = Document(stegoDocPath)

    # Access the document's styles
    styles = document.styles

    # Create and add a new, stego character style
    # It will be used as the default character style
    # Without replacing any set styles
    base_style = styles.add_style('stego_style', WD_STYLE_TYPE.CHARACTER)
    new_style = cast(CharacterStyle, base_style) # the same as CharacterStyle(base_style)

    # Defining the stego-style's font theme, size and colour
    new_style_font = new_style.font
    new_style_font.name = 'Britannic Bold'
    new_style_font.size = Pt(16)
    new_style_font.color.rgb = RGBColor(0x23, 0x23, 0x23) # Black

    # This applies stego-style to the default character style only
    print(f"Changing the default character style format of: {Path(stegoDocPath).name}")
    # Loop through each paragraph
    for paragraph in document.paragraphs:

        # Get each paragraph element's properties
        paragraph_element = paragraph._p
        paragraph_properties = paragraph_element.pPr

        # Check if there is a style on a paragraph level
        if paragraph_properties != None:
            pStyle = paragraph_properties.pStyle
        else:
            pStyle = None

        # Loop through each run
        for run in paragraph.runs:
            # Get each run element's properties
            run_element = run._r
            run_properties = run_element.rPr

            # Check if there is a style on a run level
            if run_properties != None:
                rStyle = run_properties.rStyle
            else:
                rStyle = None

            # If there is no style on either level
            # Then the text uses the default style
            # So apply the stego-style on the text
            if pStyle == None and rStyle == None:
                run.style = new_style

    # Save
    document.save(attackedStegoDocPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("04_format_attack", formatting_attack, False)