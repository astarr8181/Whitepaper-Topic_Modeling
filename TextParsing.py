#!/usr/bin/env python

import PyPDF2
import os
import docx
import pandas as pd


#parses the text from PDFs and .docx files in a directory and exports them to a .csv
#check the FileContent column of the csv to identify any errors and the file associated with them
def main():
    #sets the path to the current directory. Run the script in the directory containing the files to be parsed or reset the path to the file location
    folder = './'
    
    #creates a list of all .docx files in the directory
    word_docs = []
    word_docs += [file for file in os.listdir(folder) if file.endswith('.docx')]
    
    #creates a list of all .pdf files in the directory
    pdfs = []
    pdfs += [file for file in os.listdir(folder) if file.endswith('.pdf')]
    
    text =[]
    for document in word_docs:
        text += [reading_word_documents(document)]
        
    pdf_text = []
    for pdf in pdfs:
        try:
            pdf_text += [reading_pdfs(pdf)]
        except:
            pdf_text += ['Error Reading PDF']
            
    content = text + pdf_text
    
    filenames = word_docs + pdfs
    df = pd.DataFrame(
        {'FileName': filenames,
         'FileContent': content
        })
    
    #outputs the file name and its content to a row in a csv. If there is no content and error message is inserted
    output = df.fillna('Error Reading Text')
    output = pd.DataFrame(output)
    output.to_csv('whitepapers_text.csv')


#function parses text from docx files by paragraph and aggregates it
def reading_word_documents(file_name):
    doc = docx.Document(file_name)
    
    completed_text = []
    
    for paragraph in doc.paragraphs:
        completed_text.append(paragraph.text)
        
    return '\n' .join(completed_text)


#function extracts text from PDFs
def reading_pdfs(file_name):
    read_pdf = PyPDF2.PdfFileReader(file_name)
    page = read_pdf.getPage(0)
    page_content = page.extractText()
    
    return page_content



if __name__ == "__main__":
    main()