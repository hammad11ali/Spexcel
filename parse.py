
import pyexcel
import sys
import pyexcel_xlsx
import time
import os

os.system('cls')

print("\n\n***Welcome to Spool to Excel Parser and Converter***\n\n")

if len(sys.argv) < 2:
    print('Please provide spool file name with extension')
    print('Format: Parser filename')
    sys.exit()
try:
    file = open(sys.argv[1], 'r')
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
once = False
read = False
while True:
    line = file.readline()
    if not line:
        break
    if line.startswith('SQL>'):
        new = True
        data = False
        if rows:
            rows.remove(rows[0])
        final.append((query, header, rows))
        query = ""
        header = ""
        rows = []
    if new:
        query += ' '+line
        if line == "\n":
            new = False
            data = True
            once = False
            continue
        continue
    if data:
        if line == "\n" and once == False:
            once = True
            continue
        if once and line == "\n":
            continue
        if header == "":
            header = line
        else:
            if not line.count(" selected") and not line.count("old:") and not line.count("new:"):
                rows.append(line)
if rows:
    rows.remove(rows[0])
final.append((query, header, rows))
final.remove(final[0])
file.close()

print('Data Parsing complete')
print('Creating Excel Sheets')

sheet1 = []
sheets = []
for statement in final:
    query, header, rows = statement
    if len(rows) == 0:
        sheet1.append([query])
        sheet1.append([header])
        sheet1.append([])
        continue
    sheet = [query.upper()]
    header = [value.strip() for value in header.split(';')]
    sheet.append(header)
    for row in rows:
        row = [value.strip() for value in row.split(';')]
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
print('Saving to file '+filename)
filename = os.getcwd()+'\\'+filename
pyexcel.save_book_as(bookdict=workbook, dest_file_name=filename)
print('Full Path to file: '+filename)
print('Saved Successfully')
