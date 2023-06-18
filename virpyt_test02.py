#import openpyxl
import VirPyT
from VirPyT import Workbook
    


wb = VirPyT.Workbook('sample.xlsx')
print(wb.workbook)



for sheet in wb.worksheets:
    print("Found sheet named %s" %sheet)
    
    #for table in sheet.tables():
     #       print("Found table named %t" %table)

     #   for column_header in table.header:  # iterate through cells that make
                                            # up the header
      #  print(“col header is %s” % column_header.value)
       # for row in table:
        #    cell = row['<col name>']        # we want to be able to access a
                                            # cell according to the header name.
         #   print('row info is %s' % cell.value)
