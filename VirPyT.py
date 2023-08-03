# use openpyxl to open an excel sheet
# use python classes to sort attributes of the excel sheet


import openpyxl
        

# workbook class with pointer to the workbook
class VirpytWorkbook():

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

    
    def save(self):
        self._workbook.save()
        
        

# sheet class with pointers to sheets in the workbook        
class Sheet():
    
    def __init__(self, sheet, name):
        self._sheet = sheet     #openpyxl sheet obj
        self._name = name
        self._tables = {}       #{startcell : values}
        

    @property
    def name(self):
        return self._name

    #table objects
    @property
    def tables(self):
        self.find_tables()
        return [v for _,v in self._tables.items()]
        

        
    def find_tables(self):

        tablestartrow = self._sheet.min_row
        tablestartcol = self._sheet.min_column
        
        #find first cell with values using min_row, min_column
        startcell = self._sheet.cell(row=self._sheet.min_row,
                                     column=self._sheet.min_column).coordinate

        while tablestartrow < self._sheet.max_row:
            # 0-based numrow & numcol
            # dimensions of table
            numrow = 0
            numcol = 0

            # scan until empty column is found
            for col in self._sheet.iter_cols(min_col=tablestartcol,
                                             min_row=tablestartrow):
                if col[0].value:    #header row should extend over whole table
                    numcol += 1     #count columns in header row
                else:
                    break           #break when no value in header row

            # scan until empty row is found
            for row in self._sheet.iter_rows(min_row=tablestartrow,
                                             min_col=tablestartcol,
                                             max_col=tablestartcol+numcol-1):
                #sometimes theres a gap where one column doesnt have data
                emptyrow = True
                for cell in row:
                    if cell.value:
                        emptyrow = False
                if emptyrow:
                    break
                else:
                    numrow += 1

            coords = (tablestartcol, tablestartrow)
            
            tableendcol = (tablestartcol + numcol -1)
            tableendrow = (tablestartrow + numrow -1)

            self._tables[startcell] = Table(self._sheet, coords,
                                            numcol, numrow)


            # find next table
            startcell, tablestartcol, tablestartrow = self.startcell(
                                                      tableendcol,
                                                      tableendrow,
                                                      numcol,
                                                      numrow)


    def startcell(self, tableendcol, tableendrow, numcol, numrow):
        
        # search vertically for next table
        #assume tables are aligned, check first col for empty cells
        for row in self._sheet.iter_rows(min_row=tableendrow):
            if not row[self._sheet.min_row].value:
                tableendrow +=1
        
        
        '''
        # start seraching horizontally if maxrows reached
        if tableendrow == self._sheet.max_row:
            for col in self._sheet.iter_cols(min_col = tableendcol):
                if not col[0].value:
                    tableendcol += 1

            return self._sheet.cell(row=(tableendrow-numrow+1),
                                    column=tableendcol).coordinate
        else:
            return self._sheet.cell(row=tableendrow,
                                    column=(tableendcol-numcol+1)).coordinate
        '''
        
        startcell = self._sheet.cell(column=(tableendcol-numcol+1),
                                     row=tableendrow)
        return startcell, startcell.column, startcell.row

    


        #scan til valued cell, use header as startcell
        
        
        
        

# table class to scan empty cells that bound the table
# ((or look for cell border formatting in the file))
class Table():
    def __init__(self, sheet, coords, numcol, numrow):
        
        #defines table object with starting cell and dimensions
        self._sheet = sheet
        
        self._coords = coords
        self._numcol = numcol
        self._numrow = numrow
        self._rows = []
        self._header = {}

        
    @property
    def header(self):
        
        return Row(self, 0)

    @property
    def rows(self):
        return [Row(self, index) for index in range(self._numrow)]
        

    @property
    def columns(self):
        columns = []
        idx = 0
        while idx < len(self.table[0]):
            columns.append([row[idx].value for row in self.table])
            idx+=1
        return columns


    @property
    def coords(self):
        return self._coords
    


class Row():
    
    def __init__(self, table, index):
        self._table = table
        self._index = index

    def rowvals(self):
        sheet = self._table._sheet
        
        y = index+ self._table.coords[1]
        
        retval = []
        for x in range(self._table._coords[0],
                       self._table.numcol + self._table._coords[0]):
            cellval = sheet.cell(x,y).value
            retval.append(cellval)

        return retval
        

    def __getitem__(self, key):
        pass




class Column():
    
    def __init__(self, table, index):
        self._table = table
        self._index = index

    def rowvals(self):
        sheet = self._table._sheet
        
        y = index+ self._table.coords[1]
        
        retval = []
        for x in range(self._table._coords[0],
                       self._table.numcol + self._table._coords[0]):
            cellval = sheet.cell(x,y).value
            retval.append(cellval)

        return retval
        

    def __getitem__(self, key):
        return self._row[key]

            
#class Cell():

              

if __name__ == '__main__':
    #workbookname = input(print("Input workbookname: "))
    
    wb = VirpytWorkbook('sample.xlsx')
    print(wb)        #address of openpyxl workbook object

    print(wb.sheets)        #list of sheet object addresses
    
    print(wb.sheetnames)    #list of names of worksheets

    for sheet in wb.sheets:     #prints each sheet title
        print("Found sheet named %s" % sheet.name)
        for table in sheet.tables:
            print("Found table: ", table.coords,
                  table._numcol, table._numrow)

            
