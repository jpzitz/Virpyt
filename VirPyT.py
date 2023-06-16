import openpyxl
from openpyxl import Workbook as book
        

# workbook class with pointer to the workbook
class Workbook():
    def __init__():
        file = openpyxl.workbook()
            
    def __init__(self, file):
        self.file = file
            
    @property
    def file(self):
        """The file property."""
        print("Get file")
        return self._file

    @file.setter
    def file(self, filename):
        print("Set filename")
        self._file = filename

        
    def sheets(self):
        return [Sheet.temp_sheet for temp_sheet in self._wb.worksheets] 


# sheet class with pointers to sheets in the workbook        
class Sheet():
    def __init__(self, sheet):
        self._sheet = [openpyxl.workbook.sheetnames]
        



        
#class Table():
#class Row():
#class Cell():
      
#worksheet = workbook.active
#print(wb.sheetnames)
        

if __name__ == '__main__':
    filename = input(print("Input filename: "))
    
    wb = openpyxl.load_workbook(filename)
    

    workbook = Workbook()
    print(workbook._file)

    worksheet = wb.active
    if worksheet:
        print("ok!")

    for sheet in workbook.sheets:
        print("Found sheet named %s" %sheet.title)

