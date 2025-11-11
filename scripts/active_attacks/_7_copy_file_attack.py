from pathlib import Path
import shutil
import unified_active_attack_file

def copy_attack(
        stegoDocPath: str, 
        attackedStegoDocPath: str
        ) -> None:
    
    # Simply copy the stego-file
    print(f"Copying: {Path(stegoDocPath).name} to a different location")
    shutil.copy(stegoDocPath, attackedStegoDocPath)

if __name__ == "__main__":
    unified_active_attack_file.unified_attack("07_copy_attack", copy_attack, False)