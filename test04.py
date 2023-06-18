import openpyxl

class Book():
    def __init__(self, file):
        self.file = openpyxl.load_workbook(file)

    @property
    def worksheets(self):
        return self.file.worksheets

    

wb = Book('sample.xlsx')
print(wb.file)
print(wb.file.worksheets)    #list of sheetnames
ws = wb.worksheets[0]


# openpyxl shows spreadsheets as tuples of cells in each row
# then tuples of rows in each sheet
# I want to try to figure out how to grab a cell's location in the sheet
# by accessing its row value and index in that row tuple
for row in ws.values:
    print(row)
