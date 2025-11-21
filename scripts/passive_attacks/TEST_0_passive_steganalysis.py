import _1_passive_attack_hide_in_text
import _4_passive_attack_modify_RGB_color_ch
import _5_passive_attack_unispace
import _6_passive_attack_unicode_homoglyphs

if __name__ == "__main__":

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