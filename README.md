# Steganalysis_research_for_DOCX_format

## Description
DOCX files are Office Open XML (OOXML) format word processing document files developed by Microsoft. Unlike older DOC files, which are binary files, DOCX files are zipped packages made up of several XML and multimedia files. Most of the visible content in DOCX files is saved in the package file called document.xml.

This project is focused on studying the detection of secret messages and information embedded in DOCX file content that is located in document.xml. The process of embedding secret information into files, also known as cover medium or cover objects, is called steganography, while the analysis and detection of this secret data in the files containing them, also known as stego-files, is done in a process called steganalysis.

In the project 4 rudimentary but complete steganography methods (stego-methods) for DOCX files were created based on their descriptions in several research materials:
* hide in text (hiding characters between words)
* ~multilayer hybrid~ (NOT IMPLEMENTED)
* ~two bit transformation~ (NOT IMPLEMENTED)
* modify RGB color channels (changing RGB hex values in a barely noticeable way to encode bits)
* unispace (using unicode whitespaces in various combinations to represent data)
* unicode homoglyph (utilizing similarly looking characters to represent bits)

In the project, steganalysis is performed in three parts:
1. Active steganalysis. Modifying the DOCX file in various ways in the hope of destroying or at least corrupting the secret message. In the project 10 active attacks were performed on each stego-DOCX file:
    * Insert attack. Insert a word after nth (10th) word
    * Delete attack. Delete every nth (10th) word
    * Edit/modify attack. Change every nth (10th) word
    * Formatting attack. Change the text font format
    * Impersonation attack. Save the DOCX as PDF and back
    * Save As attack. Save the DOCX file anew
    * Copy attack. Copy the DOCX file
    * Retype attack. Retype all DOCX file content to a new DOCX file
    * Clear Format attack. Use the MS Word 'Clear All Formatting' function on the entire DOCX file content
    * Document Inspect attack. Use every option in the MS Word Document Inspect tool on the DOCX file
2. Attacked file analysis. Scanning the DOCX file using the stego-method extraction algorithms to extract the secret message. 5 levels of corruption were classified using Ratcliff/Obershelp pattern recognition algorithm:
    * Safe (completely unaffected)
    * Almost safe (still very readable, only up 5% missing original content)
    * Significantly corrupted (barely readable, up to 50% missing original content)
    * Heavily corrupted (unreadable, more than 50% missing original content)
    * Missing (completely destroyed)
3. Statistic analysis. Scanning and extracting file features that can help indicate whether the DOCX file contains any unusual data, as well as analyzing the effects of the stego-methods on file make-up. Analysis was performed from 7 categories:
    * File Size
    * Element Count (w:p, w:r, w:t amount)
    * Font Size (w:sz values in w:r)
    * RGB Value (w:color values in w:r)
    * Single Char (w:t amount with a single char)
    * Space Char (whitespace amount in w:body)
    * Unicode (non-ASCII unicode symbol amount in w:body)

The code for the steganography and steganalysis is located in the scripts folder.

The data set used was 21 DOCX files (20 from this [Dataset for Doc & Docx](https://www.kaggle.com/datasets/manisha717/dataset-for-doc-and-docx), 1 created during the testing of the code). The 21 DOCX files were scanned for validity of embedding of the secret message. Those that were valid had 178 byte sized stego-message embedded into them and each file was saved. Each stego-DOCX file was then attacked using the active attacks and saved. The attacked stego-DOCX files were then scanned for the stego-message. The data set for the clean, TEST_0, stego- and attacked DOCX files is located in the data_set folder.

The results from the steganalysis are located in the results folder, as complete XLSX files and as individual CSV files.

## Requirements
The project was written entirely in Python (version 3.13.7) and the following Python libraries were used in making this project:
* Third party:
    * lxml 6.0.0
    * python-docx 1.2.0
    * pywin32 311
    * setuptools 80.9.0

* In-built:
    * os
    * pathlib
    * shutil
    * re
    * base64
    * random
    * time
    * copy
    * typing
    * collections
    * itertools
    * difflib
    * csv

The Microsoft Office is also required to be installed.

The code works only on Windows 10 and Windows 11 operating systems.

## Steps to set up
1. Have Python version 3.13 or newer installed
2. Ensure you have pip installed
3. Download the repo in any location on the PC
4. Open the terminal (CMD)
5. Navigate to the location the repo was downloaded in
6. Navigate to the root folder
7. Uninstall project data if it exists ```py -3.13 -m pip uninstall Steganalysis_research_for_DOCX_format```
8. Run this line in CMD which installs all necessary libraries ```py -3.13 -m pip install -e .```