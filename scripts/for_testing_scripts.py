from pathlib import Path
from docx import Document

docPath = str(Path("data_set/TEST_0/TEST_0.docx"))
document = Document(str(Path("data_set/TEST_0/TEST_0.docx")))
document.save(str(Path("data_set/TEST_0/TEST_1.docx")))