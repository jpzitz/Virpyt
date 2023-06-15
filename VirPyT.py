#import openpyxl
from openpyxl import Workbook
#wb = Workbook('VirPyT-tester.xlsx')



import openpyxl
wb = openpyxl.load_workbook('sample.xlsx')
ws = wb.active

#type(wb)


#    grab the active worksheet
#ws = wb.active

if ws:
    print("ok!")


s = input()

if s:
    #print(ws['A1'])

    # Data can be assigned directly to cells
    ws['A8'] = 43

    # Rows can also be appended
    #ws.append([1, 2, 3])

    # Python types will automatically be converted
    import datetime
    ws['B8'] = datetime.datetime.now()

    # Save the file
    wb.save("sample.xlsx")
else:
    print(ws['A1'].value)   #test readable
