# Project DocMerge
# Author: Akshat Johari

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

def main():
    print("Welcome to the document creation chamber.\nThis program creates .docx files with images and text, in tabular format.\n")
    name = input("Please enter the name you want this document to be given, without the .docx extension.\n")
    create_doc(name)
    
def create_doc(name):
    doc = Document()
    title = input("Please enter the title of your document.\n")
    p = doc.add_paragraph(title)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.style = doc.styles['Heading 1']
    p.style.font.size = Pt(16)
    p.style.font.bold = True
    
    sub = input("Does your document have a subheading?\nY for Yes, N for No\n")
    if sub in ("y", "Y"):
        sub = input("Please enter your subheading.\n")
        p = doc.add_paragraph(sub)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.style = doc.styles['Heading 2']
        p.style.font.size = Pt(14)
        p.style.paragraph_format.space_after = Pt(12)   
    
    doc.save(name + ".docx")
    print("File " + name + " has been created in the current directory. Fare thee well.")
          

if __name__ == "__main__":
    main()