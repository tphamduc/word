from docx import Document
from docx.shared import Pt
from docx.text.run import Font

def ReadText(filename):
    document = Document(filename)
    text = []
    for data in document.paragraphs:
        text.append(data.text)

    for x in text:
        print(x)

# ReadText("test.docx")

def FontSize(filename):
    document = Document(filename)
    all_12pt_text = []
    for p in document.paragraphs:
        for run in p.runs:
            if run.font.size == Pt(14):
                all_12pt_text.append(p.text)

    all_12pt_text_set = []
    all_12pt_text_set = set(all_12pt_text)

    for c in all_12pt_text_set:
        print(c)

# FontSize("test.docx")

# document = Document("test.docx")
# for p in document.paragraphs:
#     for run in p.runs:
#         print(run.font.size)

def FontBold(filename):
    document = Document(filename)
    all_bold = []
    for p in document.paragraphs:
        for run in p.runs:
            if run.font.bold == Font.cs_bold:
                all_bold.append(p.text)
                print(p)

    all_bold_set = []
    # all_bold_set = set(all_bold)
    #
    # for c in all_bold_set:
    #     print(c)


FontBold("test.docx")



