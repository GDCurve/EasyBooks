from openpyxl import Workbook
from openpyxl import load_workbook


def startScreen():
    print("""Welcome, please choose an option:\n
        edit | read | product | help | exit""")
    Choice = input('Choice >>> ').lower()

    if Choice == "edit":
        edit()
    elif Choice == "read":
        read()
    elif Choice == "product":
        product()
    elif Choice == "help":
        print(r"""--------------------------- HELP -----------------------------
edit => allows for addition or subtraction of product count.

read => reads the product count in warehouse

product => adds a new product to the database

help => opens help menu

exit => exits the application
---------------------------------------------------------------""")
        startScreen()
    elif Choice == "exit":
        print('exiting...')
    else:
        print('unknown command, try again')
        startScreen()

def edit():
    print('edit')

def read():
    print('read')

def product():
    print(r""" choose an option:
        list | add | remove""")
    choice = input("Choice >>> ").lower()
    if choice == "list":
        product_list()
    elif choice == "add":
        product_add()
    elif choice == "remove":
        product_remove()
    else:
        print("unknown choice, try again")
        product()


# product functions
def product_list():
    Book = load_workbook("Data.xlsx")
    Sheet = Book['Sheet']
    i = 0
    for row in Sheet:
        i = i + 1
        I = str(i)
        print(str(Sheet['A' + I].value) + "   |   " + str(Sheet['B' + I].value) + "   |   " + str(Sheet['C' + I].value))
        print("""-------------------------------------------------------------------""")

    ans = input('Continue? Y >>> ').lower()
    if ans == "y":
        startScreen()
    else:
        print('Unknown command')


def product_add():
    Book = load_workbook("Data.xlsx")
    Sheet = Book['Sheet']

    name = input("Product name >>> ")
    qty = input("Current quantity >>> ")
    max = Sheet.max_row + 1
    DeletedRows = Sheet.cell(row=1, column=999).value
    print(DeletedRows)

    # if NextID == None:
    #     NextID = 1
    if DeletedRows == 0:
        NextID = max-1
    else:
        NextID = Sheet.cell(row=max - DeletedRows, column=1).value
    print(NextID)

    Sheet.cell(row=max, column=1).value = NextID + 1
    Sheet.cell(row=max, column=2).value = name
    Sheet.cell(row=max, column=3).value = int(qty)
    Book.save('Data.xlsx')

    print(name + " was added with a quantity of " + str(qty) + " and an ID of " + str(NextID))
    ans = input('Add more? Y/N >>> ').lower()
    if ans == "y":
        product_add()
    elif ans == "n":
        startScreen()
    else:
        print('Unknown command')


def product_remove():
    Book = load_workbook("Data.xlsx")
    Sheet = Book['Sheet']
    row = input('Input ID of product >>> ')

    Sheet.delete_rows(int(row)+1)
    Sheet.cell(row=1, column=999).value = Sheet.cell(row=1, column=999).value + 1
    Book.save('Data.xlsx')

    ans = input('Remove more? Y/N >>> ').lower()
    if ans == "y":
        product_remove()
    elif ans == "n":
        startScreen()
    else:
        print('Unknown command')
