import sys
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.layout import LAParams
from cStringIO import StringIO
import xlsxwriter
def pdfparser(data):

    fp = file(data, 'rb')
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    # Create a PDF interpreter object.
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    # Process each page contained in the document.

    for page in PDFPage.get_pages(fp):
        interpreter.process_page(page)
        data =  retstr.getvalue()

    return data
def no_of_x_in_y(x, y):
    number = 0
    for i in y:
        if i == x:
            number = number + 1

    return number


def if_email(case):
    index_at_the_rate = 0
    index_dot = 0
    for i in range(0, len(case)):
        if case[i] == "@":
            index_at_the_rate = i
            break
    if no_of_x_in_y(".", case[index_at_the_rate:len(case)]) == 0:
        return 0
    if(len(case[0:index_at_the_rate])==0):
        return 0

    if no_of_x_in_y("@", case) != 1:
        return 0
    else:
        ass=1
        while ass==1:
            if(case.endswith(' ')):
                case = case[:-1]
            else:
                ass = 2
        if( case.endswith('@')):
            return 0
        else:
            return 1
from os import path
from glob import glob
def find_ext(dr, ext, ig_case=False):
    if ig_case:
        ext =  "".join(["[{}]".format(ch + ch.swapcase()) for ch in ext])
    return glob(path.join(dr, "*." + ext))

if __name__ == '__main__':
    pdfs = find_ext(".","pdf",True)
    emails = ['email','soorajceo@gmail.com']
    workbook = xlsxwriter.Workbook('Extracted.xlsx')
    worksheet = workbook.add_worksheet()
    row = 0
    col = 0
    for pdf in pdfs:
        data = pdfparser(pdf)
        given_text = data
        a = given_text.split()
        for i in range(0, len(a)):
            present_case = a[i]
            if (if_email(present_case) == 1):
                ass=1
                while ass==1:
                    if(present_case.endswith('.') or present_case.endswith(':') or present_case.endswith(';') or present_case.endswith(',') or present_case.endswith(')') or present_case.endswith(' ')):
                        present_case = present_case[:-1]
                    else:
                        ass=2
                
                print present_case
                test = ""
                for presents in present_case:
                    test = test+presents
                worksheet.write(row, col,test)
                row = row + 1
    workbook.close()
