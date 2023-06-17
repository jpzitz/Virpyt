# use openpyxl to open an excel sheet
# use python classes to sort attributes of the excel sheet


import openpyxl
        

# workbook class with pointer to the workbook
class Workbook():

    # constructor/attribute __init__() that takes a filename
    # and then uses openpyxl to open that file and store the
    # workbook pointer in a local member.  
    def __init__(self, file):
        self.file = openpyxl.load_workbook(file)

        
    '''        
    @property
    def file(self):
        """The file property."""
        print("Get file")
        return self._file


    @file.setter
    def file(self, filename):
        print("Set filename")
        self._file = filename
    '''
    @property
    def worksheets(self):
        return self.file.sheetnames

    # attribute/method that returns a list of VirPyTSheet
    #@property    
    def sheets(self):

        # constructor should be generated by
        # wrapping individual sheet returned by wb.worksheets
        return [Sheet(temp_sheet) for temp_sheet in self._file]



# sheet class with pointers to sheets in the workbook        
class Sheet():
    def __init__(self, sheet):
        self._sheet = Workbook.worksheets
        
        
#class Table():
#class Row():
#class Cell():

              

if __name__ == '__main__':
    #filename = input(print("Input filename: "))
    
    #wb = openpyxl.load_workbook('sample.xlsx')

    workbook = Workbook('sample.xlsx')
    print(workbook.file)
    
    print(workbook.sheets)
    print(workbook.worksheets)
    
    #worksheet = Sheet(workbook.file.active)
    #if worksheet:
    #    print("ok!")

    for sheet in workbook.worksheets:
        print("Found sheet named ", sheet.title)

