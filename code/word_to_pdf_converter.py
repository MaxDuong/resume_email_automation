import sys
from docx2pdf import convert
from docx import Document
import re 
import os


# ref: https://www.quora.com/How-can-I-find-and-replace-text-in-a-Word-document-using-Python
def docx_replace_regex(doc_obj, regex , replace): 
    for p in doc_obj.paragraphs: 
        if regex.search(p.text): 
            inline = p.runs 
            # Loop added to work with runs (strings with same style) 
            for i in range(len(inline)): 
                if regex.search(inline[i].text): 
                    text = regex.sub(replace, inline[i].text) 
                    inline[i].text = text 
 
    for table in doc_obj.tables: 
        for row in table.rows: 
            for cell in row.cells: 
                docx_replace_regex(cell, regex , replace)

def replace_company_name(company_name):
    
    regex1 = re.compile(r"company_name") 
    replace1 = company_name 
    filename = "./CoverLetter-Max.docx"

    doc = Document(filename) 
    docx_replace_regex(doc, regex1 , replace1) 
    doc.save('./CoverLetter-Max.docx') 

def replace_company_name_original(company_name):
    
    regex1 = re.compile(company_name) 
    replace1 = "company_name"
    filename = "./CoverLetter-Max.docx"

    doc = Document(filename) 
    docx_replace_regex(doc, regex1 , replace1) 
    doc.save('./CoverLetter-Max.docx') 
    pass

def doc_to_pdf_converter():
    # Current directory
    doc_path = os.getcwd()+'\CoverLetter-Max.docx'
    # Parent directory
    pdf_path = os.path.abspath(os.getcwd() + "/../") + '\CoverLetter-Max.pdf'
    
    # print(doc_path)
    # print(pdf_path)
    convert(doc_path, pdf_path)

def main():
    company_name = input("Enter your company's name: ")
    print(company_name)

    replace_company_name(company_name)
    doc_to_pdf_converter()
    replace_company_name_original(company_name)

if __name__ == "__main__":
    main()