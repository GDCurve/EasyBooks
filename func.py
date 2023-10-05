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
    Sheet = Book.active
    for row in Sheet:
        print()

def product_add():
    print(1)

def product_remove():
    print(1)