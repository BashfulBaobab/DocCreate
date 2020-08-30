# Project DocMerge
# Author: Akshat Johari

from docx import Document

def main():
    print("Welcome to the document creation chamber.\nThis program creates .docx files with images and text, in tabular format.\n")
    name = input("Please enter the name you want this document to be given, without the .docx extension.\n")
    create_doc(name)
    
def create_doc(name):
    doc = Document()
    title = input("Please enter the title of your document.\n")
    p = doc.add_paragraph(title)
    doc.save(name + ".docx")
    

if __name__ == "__main__":
    main()