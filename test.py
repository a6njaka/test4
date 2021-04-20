import fitz
# https://pymupdf.readthedocs.io/en/latest/textpage/

doc = fitz.open(r"C:\Users\NJAKA\Desktop\Jack_PDF\Example.pdf")

page_number = doc.pageCount
page = doc[0]

for p in range(page_number):
    for text in doc[p].getText("blocks"):
        print(text)

