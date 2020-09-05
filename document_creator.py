# Project DocMerge
# Author: Akshat Johari

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from os import listdir
from os.path import isfile, join

def main():
    print("Welcome to the document creation chamber.\nThis program creates .docx files with images and text, in tabular format.\n")
    name = input("Please enter the name you want this document to be given, without the .docx extension.\n")
    doc = create_doc(name)
    images = img_list()
    
    doc.save(name + ".docx")
    print("File " + name + " has been created in the current directory. Fare thee well.")
    
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
    
    return doc
              
def img_list():
    data_entry = input("How would you like to enter the images?\n1. Enter manually, one at a time\n2. Enter directory with all images (note that only jpg and png images are supported)\n3. Enter txt file with all images in separate lines\nPlease enter 1, 2, or 3: \t")
    while data_entry not in ("1", "2", "3"):
        print("Invalid entry. Please try again.")
        data_entry = input("How would you like to enter the images?\n1. Enter manually, one at a time\n2. Enter directory with all images (note that only jpg and png images are supported)\n3. Enter a .txt file with all images in separate lines\nPlease enter 1, 2, or 3: \t")
    #accept manual entries
    if data_entry == "1":
        l = []
        x = input("Enter image locations, either absolute location, or relative to current working directory. If you wish to stop data entry, type esc. Please make sure the locations are correct.\n")
        while x != "esc":
            l.append(x)
            x = input("Enter next image:\n")
        return l
    
    #accept directory with all images
    elif data_entry == "2":
        x = (".png", ".jpg")
        mypath = input("Enter directory:\n")
        l = [f for f in listdir(mypath) if (isfile(join(mypath, f)) and (f.endswith(".png") or f.endswith(".JPG")))]
        return l

if __name__ == "__main__":
    main()