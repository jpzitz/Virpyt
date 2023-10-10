"""Example code to test VirPyT snakeskin wrapper for .xlsx and .csv files.
Uses sample data taken from https://www.briandunning.com/sample-data/
Accessed on July 09 2023.
"""

from Virpyt import VirpytWorkbook



def view_rows(table):
    """opens employee records from last_name attribute"""
    print("View Rows")
    lastnamecol = table.columns[table.header.index('last_name')]
    print(lastnamecol.colvals[1:])
    lastname = input(
        "To view an employee's information, please enter Last Name: ")
    try:
        employeefile = table.rows[lastnamecol.colvals.index(lastname)].rowvals
        for item in employeefile:
            print("%13s" %table.header[employeefile.index(item)], ":", item)
    except ValueError:
        print("This employee is not in the database. Please try again.")



def view_columns(table):
    """opens all of a single attribute"""
    print("View Columns")
    for item in table.header:
        print(item)
    viewcolumn = input("Please select information to view: ")

    for item in table.columns[table.header.index(viewcolumn)].colvals:
        print(item)



def menu(table):
    """menu for user selection"""
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
        print("That is not a valid choice, please try again.")
        choice = input("Select one of the command numbers above: ")



def main():
    """Main method opens the file and stores the pointer.
    Using a database of 500 imaginary employees, database includes the
    following:
        first_name, last_name, company_name, address, city, county, state, zip,
        phone1, phone2, email, web
    We can go through the database and pull out individual lists, either by row
    (an individual employee's information) or by column (a list of all of a
    certain attibute, eg. all emails)
    """

    #open the speadsheet file using openpyxl through VirPyT
    wb = VirpytWorkbook('us-500.xlsx')

    for sheet in wb.sheets:     #prints each sheet title
        print("Found sheet named %s" % sheet.name)

    # open a sheet to view tables
    sheetname = input("Please enter a sheetname to view: ")
    try:
        ws = wb.sheets[wb.sheetnames.index(sheetname)]
        print(f"Now viewing {sheetname}.")
    except ValueError:
        print("This worksheet is not in the file. Please try again.")
    except TypeError:
        print("Type Error. Please enter the sheetname again.")


    # See all tables in a worksheet, ordered by starting cell location.
    for table in ws.tables:
        print("Found table starting on ", ws._sheet.cell(table._coords[0],
                                                         table._coords[1]
                                                         ).coordinate)


    # Open a table to view contents
    tablename = input("Please enter a starting location to view a table: ")
    try:
        table = ws._tables[f"{tablename}"]
        print(f"Now viewing table starting on{tablename}")
        menu(table)
    except ValueError:
        print("This table is not in the worksheet. Please try again.")
    except TypeError:
        print("Type Error. Please enter the starting location again.")
    except KeyError:
        print("Key Error. Please enter the starting location again.")


if __name__ == '__main__':
    main()
