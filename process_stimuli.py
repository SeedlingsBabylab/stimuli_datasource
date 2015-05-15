import xlrd
import sys
import csv

from Tkinter import *
import tkFileDialog

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
    #print "row: " + str(row) + "  template value: " + str(template_sheet.cell_value(row, 12)) + "  order: " + str(order)
    row_data = []
    for col in range(13):
        val = template_sheet.cell_value(row, col)
        if template_sheet.cell_type(row, col) == 2:
            row_data.append(int(val))
        else:
            row_data.append(val)
    data.append(row_data)
    #print "_dimnrows: " + str(template_sheet._dimnrows)

    if row >= template_sheet._dimnrows - 1:
        break
    else:
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


print "data: " + str(data)


# with open(sys.argv[3], 'wb') as f:
#     for row in data:
#         first = True
#         for item in row:
#             if not first:
#                 f.write('\t')
#                 first = False
#             f.write(str(item))
#         f.write('\n')


# This replaces the commented section above.
with open(sys.argv[3], 'w') as file:

    csvwriter = csv.writer(file, delimiter='\t')

    csvwriter.writerow(data[0])     # write the header row

    for row in data[1:]:    # write each subsequent row (skipping the header)
        csvwriter.writerow(row)


class MainWindow:

    def __init__(self, master):

        self.root = master
        root.title("Stimuli Datasource")
        self.root.geometry("500x300")
        self.main_frame = Frame(root)
        self.main_frame.pack()


        self.template_file = None
        self.stimuli_file = None
        self.export_file = None


        self.template_book = None
        self.template_sheet = None

        self.stimuli_book = None
        self.stimuli_sheet = None

        self.data = []

        self.load_template_button = Button(self.main_frame,
                                           text="Load Template",
                                           command=self.load_template)

        self.load_template_button.grid(row=2,column=1)


        self.load_stimuli_button = Button(self.main_frame,
                                          text="Load Stimuli",
                                          command=self.load_stimuli)

        self.load_stimuli_button.grid(row=2,column=2)


        self.export_button = Button(self.main_frame,
                                          text="Export",
                                          command=self.export)

        self.export_button.grid(row=2,column=3)



        self.template_loaded_label = Label(self.main_frame, text="Template Loaded", fg="blue")
        self.stimuli_loaded_label = Label(self.main_frame, text="Stimuli Loaded", fg="green")


    def load_template(self):

        self.template_file = tkFileDialog.askopenfilename()

        self.template_book = xlrd.open_workbook(self.template_file)
        self.template_sheet = template_book.sheet_by_index(0)

        self.template_loaded_label.grid(row=3, column=1)

    def load_stimuli(self):

        self.stimuli_file = tkFileDialog.askopenfilename()

        self.stimuli_book = xlrd.open_workbook(self.stimuli_file)
        self.stimuli_sheet = stimuli_book.sheet_by_index(0)

        self.stimuli_loaded_label.grid(row=3, column=2)

    def run(self):



        #0) keep the header row, ...
        hdr = []
        for c in range(13):
            hdr.append(template_sheet.cell_value(0, c))
        self.data.append(hdr)
        print hdr
        #get column G's value
        order = stimuli_sheet.cell_value(1, 6)

        #find the first row corresponding to the G subset
        row = 1
        while True:
            if self.template_sheet.cell_value(row, 12) == order:
                break
            row += 1

        #0) ..., and then take the subset of the template that is the order corresponding to column G's value (1-4)
        while self.template_sheet.cell_value(row, 12) == order:
            #print "row: " + str(row) + "  template value: " + str(template_sheet.cell_value(row, 12)) + "  order: " + str(order)
            row_data = []
            for col in range(13):
                val = self.template_sheet.cell_value(row, col)
                if self.template_sheet.cell_type(row, col) == 2:
                    row_data.append(int(val))
                else:
                    row_data.append(val)
            self.data.append(row_data)
            #print "_dimnrows: " + str(template_sheet._dimnrows)

            if row >= self.template_sheet._dimnrows - 1:
                break
            else:
                row += 1

        #1) Use the 20 rows (after the header row) in columns A and B to write into columns B through D of the spreadsheet
        #3) replace 1-16 in the .wav and .jpg with the words numbered 1-16 (e.g. 1.jpg becomes banana.jpg and can_banana.jpg)
        #4) IF there is something in column e that is not NA, replace with that instead of with the word in column B (e.g. sock3 instead of sock) ONLY in columns B&C not in column D
        for r in range(5, len(self.data)):
            index = int(self.data[r][1].split('.')[0])
            col_e = self.stimuli_sheet.cell_value(index + 4, 4)
            if col_e == "NA":
                data[r][1] = self.stimuli_sheet.cell_value(index + 4, 1) + ".jpg"
            else:
                data[r][1] = self.stimuli_sheet.cell_value(index + 4, 4) + ".jpg"

            index = int(self.data[r][2].split('.')[0])
            col_e = self.stimuli_sheet.cell_value(index + 4, 4)
            if col_e == "NA":
                self.data[r][2] = self.stimuli_sheet.cell_value(index + 4, 1) + ".jpg"
            else:
                self.data[r][2] = stimuli_sheet.cell_value(index + 4, 4) + ".jpg"

            prefix = self.data[r][3].split('.')[0].split('_')[0]
            index = int(self.data[r][3].split('.')[0].split('_')[1])
            self.data[r][3] = "%s_%s.wav" % (prefix, self.stimuli_sheet.cell_value(index + 4, 1))

        #5) replace A:H in column F with the pairs corresponding to A:H in column I of the stimuli spreadsheet
            self.data[r][5] = self.stimuli_sheet.cell_value(ord(self.data[r][5]) - ord('A') + 1, 8)


        #2) replace practice1.jpg-practice4.jpg with the first four words of 'stimuli' labeled p1-p4
        for r in range(1, 5):
            word = self.stimuli_sheet.cell_value(r, 1)
            self.data[r][1] = self.data[r][1].replace("practice%d" % r, word)
            self.data[r][3] = self.data[r][3].replace("practice%d" % r, word)



    def export(self):

        self.run()

        self.export_file = tkFileDialog.asksaveasfilename()

        with open(self.export_file, 'w') as file:

            csvwriter = csv.writer(file, delimiter='\t')

            csvwriter.writerow(data[0])     # write the header row

            for row in data[1:]:    # write each subsequent row (skipping the header)
                csvwriter.writerow(row)

if __name__ == "__main__":

    root = Tk()
    MainWindow(root)
    root.mainloop()


