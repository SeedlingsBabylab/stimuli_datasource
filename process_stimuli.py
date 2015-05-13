import xlrd
import sys

if len(sys.argv) != 4:
    print "missing arguments"
    print "usage: process_stimuli.py <template.xlsx> <stimuli.xslx> <output.csv>"
    sys.exit(0)

stimuli_book = xlrd.open_workbook(sys.argv[2])
stimuli_sheet = stimuli_book.sheet_by_index(0)
template_book = xlrd.open_workbook(sys.argv[1])
template_sheet = template_book.sheet_by_index(0)

data = []

#0) keep the header row, ...
hdr = []
for c in range(13):
    hdr.append(template_sheet.cell_value(0, c))
data.append(hdr)
print hdr
#get column G's value
order = stimuli_sheet.cell_value(1, 6)

#find the first row corresponding to the G subset
row = 1
while True:
    if template_sheet.cell_value(row, 12) == order:
        break
    row += 1

#0) ..., and then take the subset of the template that is the order corresponding to column G's value (1-4)
while template_sheet.cell_value(row, 12) == order:
    row_data = []
    for col in range(13):
        val = template_sheet.cell_value(row, col)
        if template_sheet.cell_type(row, col) == 2:
            row_data.append(int(val))
        else:
            row_data.append(val)
    data.append(row_data)
    row += 1

#1) Use the 20 rows (after the header row) in columns A and B to write into columns B through D of the spreadsheet
#3) replace 1-16 in the .wav and .jpg with the words numbered 1-16 (e.g. 1.jpg becomes banana.jpg and can_banana.jpg)
#4) IF there is something in column e that is not NA, replace with that instead of with the word in column B (e.g. sock3 instead of sock) ONLY in columns B&C not in column D
for r in range(5, len(data)):
    index = int(data[r][1].split('.')[0])
    col_e = stimuli_sheet.cell_value(index + 4, 4)
    if col_e == "NA":
        data[r][1] = stimuli_sheet.cell_value(index + 4, 1) + ".jpg"
    else:
        data[r][1] = stimuli_sheet.cell_value(index + 4, 4) + ".jpg"

    index = int(data[r][2].split('.')[0])
    col_e = stimuli_sheet.cell_value(index + 4, 4)
    if col_e == "NA":
        data[r][2] = stimuli_sheet.cell_value(index + 4, 1) + ".jpg"
    else:
        data[r][2] = stimuli_sheet.cell_value(index + 4, 4) + ".jpg"

    prefix = data[r][3].split('.')[0].split('_')[0]
    index = int(data[r][3].split('.')[0].split('_')[1])
    data[r][3] = "%s_%s.wav" % (prefix, stimuli_sheet.cell_value(index + 4, 1))

#5) replace A:H in column F with the pairs corresponding to A:H in column I of the stimuli spreadsheet
    data[r][5] = stimuli_sheet.cell_value(ord(data[r][5]) - ord('A') + 1, 8)


#2) replace practice1.jpg-practice4.jpg with the first four words of 'stimuli' labeled p1-p4
for r in range(1, 5):
    word = stimuli_sheet.cell_value(r, 1)
    data[r][1] = data[r][1].replace("practice%d" % r, word)
    data[r][3] = data[r][3].replace("practice%d" % r, word)

with open(sys.argv[3], 'wb') as f:
    for row in data:
        first = True
        for item in row:
            print item
        if not first:
            f.write('\t')
            first = False
        f.write(str(item))
    f.write('\n')

