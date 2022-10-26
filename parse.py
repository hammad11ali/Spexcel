
import string
import pyexcel
import sys
import pyexcel_xlsx
import time
import os

os.system('cls')

print("\n\n***Welcome to Spool to Excel Parser and Converter***\n\n")

# if len(sys.argv) < 2:
#     print('Please provide spool file name with extension')
#     print('Format: Parser filename')
#     sys.exit()
try:
    file = open(sys.argv[1], 'r')
    # file = open('test.SPL', 'r')
except:
    print('File not found')
    sys.exit()

print('File opened')
print('Reading and Parsing file data')
new = False
data = False

query = ""
header = ""
rows = []
final = []
states = ['QUERY', 'HEADER', 'DASHES', 'DATA']
state = states[0]
headings = []
row_data = []
dashes = []
endline = False
while True:
    line = file.readline()
    if not line:
        break
    if state == 'QUERY':
        if line == '\n':
            state = states[1]
            continue
        if line.startswith('SQL>'):
            query = line
            continue
        else:
            query += ' '+line
            continue
    if state == 'HEADER':
        if line.count('no rows selected.') > 0:
            state = states[0]
            continue
        # if line contain --- then header has ended
        if line.startswith('---'):
            state = states[2]
            headings = header.split(';')
            headings = [x.strip() for x in headings]
            dashes = line.split(';')
            continue
        else:
            header += ' '+line
            continue
    if state == 'DASHES':
        if line.count('no rows selected.') > 0:
            state = states[0]
            continue
        if line.startswith('---'):
            dashes = dashes[:]+line.split(';')
            if len(dashes) == len(headings):
                state = states[3]
                dashes = []
            continue
        else:
            state = states[3]
            row_data = line.split(';')
            row_data = [x.strip() for x in row_data]
            if len(row_data) == len(headings):
                rows.append(row_data)
                row_data = []
            continue
    if state == 'DATA':
        if line.count("rows selected.") > 0:
            state = states[0]
            endline = False
            final.append((query, headings, rows))
            query = ""
            header = ""
            rows = []
            headings = []
        if line == "\n" and endline == False:
            endline = True
            continue
        elif endline and line == "\n":
            state = states[0]
            endline = False
            final.append((query, headings, rows))
            query = ""
            header = ""
            rows = []
            headings = []
            continue
        else:
            row_data = line.split(';')
            row_data = [x.strip() for x in row_data]
            if len(row_data) == len(headings):
                rows.append(row_data)
                row_data = []
            continue


if query != "":
    final.append((query, headings, rows))
file.close()

print('Data Parsing complete')
print('Creating Excel Sheets')

sheet1 = []
sheets = []
for statement in final:
    query, header, rows = statement
    if len(rows) == 0:
        sheet1.append([query])
        continue
    sheet = [query.upper()]
    header = [value.strip() for value in header]
    sheet.append(header)
    for row in rows:
        sheet.append(row)
    sheets.append(sheet)

workbook = {'Sheet1': sheet1}
counter = 2
for sheet in sheets:
    name = "Sheet"+str(counter)
    if sheet[0].count("FROM"):
        name = sheet[0].split("FROM")[1].split(" ")[1]
        if name == "":
            name = "Sheet"+str(counter)
    workbook[name] = sheet[1:]
    counter += 1
filename = sys.argv[1].split('.')[0]+str(int(time.time()))+'.xlsx'
# filename = 'test'+str(int(time.time()))+'.xlsx'
print('Saving to file '+filename)
filename = os.getcwd()+'\\'+filename
pyexcel.save_book_as(bookdict=workbook, dest_file_name=filename)
print('Full Path to file: '+filename)
print('Saved Successfully')
