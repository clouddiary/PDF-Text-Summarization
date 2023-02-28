from transformers import pipeline
from spacy.lang.en.stop_words import STOP_WORDS
from heapq import nlargest
import PyPDF2
import xlsxwriter


filename="Anti bribary policy.pdf";
print("Name of file uploaded:", filename)

pdfFileObj = open(filename,'rb')

# The pdfReader variable is a readable object that will be parsed.
pdfReader = PyPDF2.PdfReader(pdfFileObj)

# Discerning the number of pages will allow us to parse through all the pages.
num_pages = len(pdfReader.pages)

#
# Create the Spread Sheet
#
workbook = xlsxwriter.Workbook('pdf_analysis.xlsx')
bold = workbook.add_format({'bold': True})
row = 0
SheetName = 'pdf_analysis'
worksheet = workbook.add_worksheet('pdf_analysis')
worksheet.write(row, 0, 'PageNumber', bold)
worksheet.write(row, 1, 'PageContent', bold)
worksheet.write(row, 2, 'PageSummary', bold)

page_content ="";

summarizer = pipeline("summarization", model="facebook/bart-large-cnn")



for page_number in range(0, num_pages):
    try:       
        row += 1
        pageObj = pdfReader.pages[page_number]
        page_content = pageObj.extract_text()    
        page_summary = summarizer(page_content, max_length=130, min_length=30)
        page_sum = page_summary[0]["summary_text"]
        print(page_sum)
        result = {  
                "page": page_number,  
                "content": page_content  
            } 
        print("Page# : " , page_number)
        print("Page Content : " , page_content)
        worksheet.write(row, 0,page_number)
        worksheet.write(row, 1,page_content)
        worksheet.write(row, 2,page_sum)
    except Exception as e:
        print(e)
        continue
       # worksheet.write(row, 0,page_number)
       # worksheet.write(row, 1,page_content)
       # worksheet.write(row, 2,e.)  
        
workbook.close()
print("Information collected & dumped into csv")

