from scripts.passive_attacks import _1_passive_attack_hide_in_text
from scripts.passive_attacks import _4_passive_attack_modify_RGB_color_ch
from scripts.passive_attacks import _5_passive_attack_unispace
from scripts.passive_attacks import _6_passive_attack_unicode_homoglyphs
from scripts.passive_attacks import unified_passive_attack_file

# Main
def main() -> None:
    desired_file = "TEST_0.docx"
    # Choose stego-method to attack
    # stego_method_1
    _1_passive_attack_hide_in_text.main(desired_file)

    # stego_method_4
    _4_passive_attack_modify_RGB_color_ch.main(desired_file)

    # stego_method_5
    _5_passive_attack_unispace.main(desired_file)

    # stego_method_6
    _6_passive_attack_unicode_homoglyphs.main(desired_file)

    unified_passive_attack_file.export_to_csv_all(desired_file.rsplit('.', 1)[0])

if __name__ == "__main__":
    main()