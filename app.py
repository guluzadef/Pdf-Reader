# Import Library
import PyPDF2
import xlsxwriter
from os import listdir
from os.path import isfile, join
path = '{your_path}'
onlyfiles = [f for f in listdir(path) if isfile(join(path, f))]
# print(onlyfiles)
workbook = xlsxwriter.Workbook('Report.xlsx')
excel = workbook.add_worksheet()
row = 0
col = 0
for e in onlyfiles:
    l = []
    content=''
    if e.split('.')[1]=='pdf':
        print(e)
        file = open(e, 'rb')
        print(file)
        reader = PyPDF2.PdfFileReader(file)
        for i in range(0, 1):
            content = reader.getPage(i).extractText() + "\n"
            content = " ".join(content.replace(u"\xa0", " ").strip().split()) 
            l.append(content)
        for items in l:
            items = items.split('/')
            excel.write(row, col, items)
            # if need diffrent rows and colums
            # excel.write(row, col+1, items)
            # row +=1

        workbook.close()


