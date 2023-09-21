# Example code to test VirPyT snakeskin wrapper for .xlsx and .csv files.
# Uses sample data taken from https://www.briandunning.com/sample-data/
# Accessed on July 09 2023.

import Virpyt
from Virpyt import VirpytWorkbook



def view_rows(table):
    print("View Rows")
    lastnamecol = table.rows[0].index('last_name')
    print(lastnamecol)
    lastname = input(
        "To view an employee's information, please enter Last Name: ")
    try:
        employeefile = table.rows[table.columns[lastnamecol].index(lastname)]
        for item in employeefile:
            print(item)


    except ValueError:
        print("This employee is not in the database. Please try again.")



def view_columns(table):
    print("View Columns")
    for item in table.rows[0]:
        print(item)
    viewcolumn = input("Please select information to view: ")

    for item in table.columns[table.rows[0].index(viewcolumn)]:
        print(item)



def menu(table):
    # menu for user selection
    print('------------------------------')
    print("{:5}".format("1"), "View Employee Information\n"+
          "{:5}".format("2"), "View List of Employees (or other information)\n"+
          "{:5}".format("99"), "Quit Program")
    print('------------------------------')

    choice = input("\nSelect one of the command numbers above: ")
    while int(choice) <= 99:
        if choice == '1':
            view_rows(table)
            menu(table)
            break
        if choice == '2':
            view_columns(table)
            menu(table)
            break
        if choice == '99':
            print("Program exiting, have a nice day.")
            break

        else:
            print("That is not a valid choice, please try again.")
            choice = input("Select one of the command numbers above: ")



def main():
    # Using a database of 500 imaginary employees,
    # database includes the following:
        # first_name, last_name, company_name,
        # address, city, county, state, zip,
        # phone1, phone2, email, web
    # We can go through the database and pull out individual lists,
    # either by row (an individual employee's information)
    # or by column (a list of all of a certain thing, eg. all emails)

    #open the speadsheet file using openpyxl through VirPyT
    wb = VirpytWorkbook('us-500.xlsx')

    for sheet in wb.sheets:     #prints each sheet title
        print("Found sheet named %s" % sheet.name)

    # open a sheet to view tables
    sheetname = input("Please enter a sheetname to view: ")
    try:
        ws = wb.sheets[wb.sheetnames.index(sheetname)]
        print("Now viewing %s" %sheetname)
    except ValueError:
        print("This worksheet is not in the file. Please try again.")
    except TypeError:
        print("Type Error. Please enter the sheetname again.")


    # See all tables in a worksheet, ordered by starting cell location.
    for table in ws.tables:
        print("Found table starting on %s" % table._startcell)


    # Open a table to view contents
    tablename = input("Please enter a starting location to view a table: ")
    try:
        table = ws._table['%s' %tablename]
        print("Now viewing table starting on %s" %tablename)
        menu(table)
    except ValueError:
        print("This table is not in the worksheet. Please try again.")
    except TypeError:
        print("Type Error. Please enter the starting location again.")


   



if __name__ == '__main__':
    main()
