# use openpyxl to open an excel sheet
# use python classes to sort attributes of the excel sheet


import openpyxl
        

# workbook class with pointer to the workbook
class Workbook():

    # constructor/attribute __init__() that takes a workbookname
    # and then uses openpyxl to open that workbook and store the
    # workbook pointer in a local member.  
    def __init__(self, filename):
        self._workbook = openpyxl.load_workbook(filename)
    

    # .sheetnames returns a list of sheetnames
    @property
    def sheetnames(self):
        return self._workbook.sheetnames


    # returns a list of VirPyTSheet
    @property
    def sheets(self):
                #wraps openpyxl sheet objects using Sheet class
        return [Sheet(self._workbook[sheetname], sheetname)
                for sheetname in self.sheetnames]

    #returns range of all value-containing cells (eg: A1:E7)
    def table_dims(self, sheetname):
        return self._workbook[sheetname].calculate_dimension()
                

    def save(self):
        self._workbook.save()
        
        



# sheet class with pointers to sheets in the workbook        
class Sheet():
    
    def __init__(self, sheet, name):
        self._sheet = sheet
        self._name = name

    @property
    def name(self):
        return self._name

    #table objects
    def tables(self):
        return self.sheet.calculate_dimension()

        
        

# table class probably to scan empty cells that bound the table
# or look for cell border formatting in the file
class Table():
    def __init__(self, table):
        self.table = table

    @property
    def range(self):
        return self.calculate_dimension()





        _startcell #set at A1 for now
        _numrow
        _numcol

        #run through sheet to find starting cell
        #work on identifying tables in a sheet
        #what do we want to do with it
        #list of headers, put in dict to get col num

        

class Row():
    def __init__(self, row ):
        self.row = row

            
#class Cell():

              

if __name__ == '__main__':
    #workbookname = input(print("Input workbookname: "))
    
    wb = Workbook('sample.xlsx')
    print(wb)        #address of openpyxl workbook object

    print(wb.sheets)        #list of sheet object addresses
    
    print(wb.sheetnames)    #list of names of worksheets

    for sheet in wb.sheets:     #prints each sheet title
        print("Found sheet named %s" % sheet.name)

    
    for ws_name in wb.sheetnames:

        print(f"worksheet name: {ws_name}")
        # open each worksheet one at a time
        #ws = wb._workbook[ws_name]
        print(wb.table_dims(ws_name))   #shows range of value-containing cells
        




    wb.save

    
    



