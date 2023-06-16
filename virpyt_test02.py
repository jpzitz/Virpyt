import openpyxl
from openpyxl import Workbook

class VirPyTWorkbook:
    def __init__(self, name):
        self.name = name
        
#class VirPyTSheet():
#class VirPyTTable():
#class VirPyTRow():
#class VirPyTCell():

wb = openpyxl.load_workbook('sample.xlsx')
  
#worksheet = workbook.active
print(wb.sheetnames)

for sheet in wb:
    print("Found sheet named %s" %sheet.title)
    
    for table in sheet.tables:
            print("Found table named %t" %table.title)

     #   for column_header in table.header:  # iterate through cells that make
                                            # up the header
      #  print(“col header is %s” % column_header.value)
       # for row in table:
        #    cell = row['<col name>']        # we want to be able to access a
                                            # cell according to the header name.
         #   print('row info is %s' % cell.value)
