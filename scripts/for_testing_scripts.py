from pathlib import Path
import docx
import csv

# base = "data_set/stego_files"
# for directories in Path(base).iterdir():
#     print(directories.name)
#     for file in directories.iterdir():
#         if file.is_file():
#             # TE: apstrāde konkrētajai apakšmapei
#             docPath = f"{base}/{directories.name}/{file.name}"
#             print("  fails:", docPath)

# if '010100000100010101010010010100110100100001001001010011100100011100100000010100110100000101001001010011000101001100100000010001100101001001001111010011010010000001001110010110010010000001011011010100100101110100100000010010100101010101001110010001010010000001001001' == '010100000100010101010010010100110100100001001001010011100100011100100000010100110100000101001001010011000101001100100000010001100101001001001111010011010010000001001110010110010010000001011011010100100101110100100000010010100101010101001110010001010010000001001001':
#     print("vienāds")
# else:
#     print("nav vienāds")


# # let's check where string is different

# how_it_should_be = '010100000100010101010010010100110100100001001001010011100100011100100000010100110100000101001001010011000101001100100000010001100101001001001111010011010010000001001110010110010010000001011011010100100101110100100000010010100101010101001110010001010010000001001001'
# how_it_is = '010100000100010100001001010011010010000100100101001101001000000101001100010010010100110001010011001000000100011001010010010011111001001000000100111001010100000101101101010010010111100010100101010101001110010001010010000001001001'

# for i in range(len(how_it_should_be)):
#     if how_it_should_be[i] == how_it_is[i]:
#         print(f"Same at index {i}: {how_it_should_be[i]}")
#     elif how_it_should_be[i] != how_it_is[i]:
#         print(f"Different at index {i}: should be {how_it_should_be[i]}, is {how_it_is[i]}")


# index: 14, 15
# deltas: (0, 1)
# from: 2, 5
# to: 2, 6
# index: 16, 17
# deltas: (0, 1)
# from: 2, 6
# to: 2, 7

# from1 = 2
# from2 = 5
# to1 = 2
# to2 = 6
# delta1 = abs(to1 - from1)
# delta2 = abs(to2 - from2)
# print(f"deltas: ({delta1}, {delta2})")

# from1 = to1
# from2 = to2
# to1 = 2
# to2 = 7
# delta1 = abs(to1 - from1)
# delta2 = abs(to2 - from2)
# print(f"deltas: ({delta1}, {delta2})")

# result_file = f"results/font_sizes/clean.csv"
# result_file_2 = f"results/font_sizes/clean.xlsx"
# data_example = [28.0, 11.0, 11.0, 11.0, 14.0, 11.0]
# i = 1
# stringexample = '' * i
# with open(result_file, "w", encoding="utf-8", newline="") as file:
#     w = csv.writer(file, delimiter=";")
#     #w.writerow([file_name, text_element_count, *font_sizes])
    
#     # w.writerow(["Document Name"])
#     # w.writerow(["Text Element Count"])
#     # w.writerow(["Font Sizes (pt)"])

#     # w.writerow(["Document Name"])
#     # w.writerow(["Text Element Count"])
#     # w.writerow(["Font Sizes (pt)"])

#     data_set_type = "clean"
#     column_headers = ["Document Name", "Text Element Count", "Font Sizes (pt)"]
#     column_data = ["text.docx", "6", *data_example]

#     w.writerows([
#         ["Document Name", "text.docx"],
#         ["Text Element Count", "6"],
#         ["Font Sizes (pt)"],
#     ])
#     for size in data_example:
#         w.writerow([
#             [stringexample, size]
#             ])

# import pandas as pd

# read_file = pd.read_csv('results/font_sizes/clean.csv', sep=';')
# read_file.to_excel('results/font_sizes/clean.xlsx', index=None, header=True)

# from openpyxl import Workbook
# wb = Workbook()
# ws = wb.active
# with open(result_file, 'r') as f:
#     for row in csv.reader(f):
#         ws.append(row)
# wb.save(result_file_2)


from docx import Document
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

# def paragraph_default_font_size_value(styles_element) -> str | None:
#     all_styles = styles_element.findall(f".//{qn('w:style')}")
#     for style in all_styles:
#         if style.get(qn('w:type')) == 'paragraph' and style.get(qn('w:default')) == '1':
#             sz = style.find(f".//{qn('w:sz')}'")
#             if sz != None:
#                 sz_value = sz.get(qn('w:val'))
#                 if sz_value != None:
#                     return sz_value
#     # If no default paragraph style found
#     return None

# def document_default_font_size_value(styles_element) -> str | None:
#     run_properties_style = styles_element.find(f".//{qn('w:docDefaults')}/{qn('w:rPrDefault')}/{qn('w:rPr')}")
#     if run_properties_style != None:
#         sz = run_properties_style.find(qn('w:sz'))
#         if sz != None:
#             sz_value = sz.get(qn('w:val'))
#             if sz_value != None:
#                 return sz_value
#     else: # If no default document style found
#         return None

# docPath = str(Path("data_set/clean_files/TEST_0.docx"))
# print(type(docPath))
# document = Document(docPath)
# styles = document.styles
# new_style = styles.add_style('stego_style', WD_STYLE_TYPE.CHARACTER)

# print(type(new_style))

# import pandas as pd
# data = {'Name': ['ANSH', 'VANSH'], 'Age': [25, 30]}
# #data = {'Name': ['Age'], 'ANSH': [25], 'VANSH': [30]}

# df = pd.DataFrame(data)
# #df.set_index(['Name', 'ANSH', 'VANSH'], inplace=True)
# print("Original DataFrame:")
# print(df)
# transposed_df = df.transpose()
# print("\nTransposed DataFrame:")
# print(transposed_df)

#data_tp = {'Name': ['Age'], 'ANSH': [25], 'VANSH': [30]}

print(26 % 7)
print(26 / 7)
print(26 // 7)

print("Testing second workstation")