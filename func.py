from openpyxl import load_workbook


def startScreen():
    print("""Welcome, please choose an option:\n
        edit | count | product | help | exit""")
    Choice = input('Choice >>> ').lower()

    if Choice == "edit":
        clearscreen()
        id = input('Input ID of product >>> ')
        edit(id)
    elif Choice == "count":
        clearscreen()
        id = input('Input ID of product >>> ')
        count()
    elif Choice == "product":
        clearscreen()
        product()
    elif Choice == "help":
        clearscreen()
        print(r"""--------------------------- HELP -----------------------------
edit => allows for addition or subtraction of product count.

count => reads the product count in warehouse

product => allows adding, removing or listing products

help => opens help menu

exit => exits the application
---------------------------------------------------------------""")
        startScreen()
    elif Choice == "exit":
        print('exiting...')
    else:
        print('unknown command, try again')
        startScreen()

def edit(id):
    Book = load_workbook("Data.xlsx")
    Sheet = Book['Sheet']

    i = 0
    for row in Sheet:
        i = i + 1
        if Book.active.cell(row=i, column=1).value == int(id):
            name = Book.active.cell(row=i, column=2).value
            count = Book.active.cell(row=i, column=3).value
            clearscreen()
            print('You are about to edit the count of ' + name + " which is stored at a quantity of " + str(count))
            ans = input('Continue? Y / N >>> ').lower()
            if ans == "y":
                op = input('Would you like to Add | Remove >>> ').lower()
                qty = int(input('How much? >>> '))
                if op == "add":
                    Book.active.cell(row=i, column=3).value = Book.active.cell(row=i, column=3).value + qty
                    Book.save('Data.xlsx')
                    startScreen()
                elif op == "remove":
                    Book.active.cell(row=i, column=3).value = Book.active.cell(row=i, column=3).value - qty
                    Book.save('Data.xlsx')
                    startScreen()
                else:
                    print("Unknown command")
            else:
                print('Cancelling')
                startScreen()


def count(id):
    Book = load_workbook("Data.xlsx")
    Sheet = Book['Sheet']
    i = 0
    for row in Sheet:
        i = i + 1
        if Book.active.cell(row=i, column=1).value == int(id):
            name = Book.active.cell(row=i, column=2).value
            qty = Book.active.cell(row=i, column=3).value
            print("there's " + str(qty) + " of " + str(name) + " stored")

    ans = input('Count another product? Y/N >>> ').lower()
    if ans == "y":
        count()
    elif ans == "n":
        startScreen()
    else:
        print('Unknown command')


def product():
    print(r""" choose an option:
        list | add | remove""")
    choice = input("Choice >>> ").lower()
    if choice == "list":
        clearscreen()
        product_list()
    elif choice == "add":
        clearscreen()
        product_add()
    elif choice == "remove":
        clearscreen()
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
        startScreen()


def product_add():
    Book = load_workbook("Data.xlsx")
    Sheet = Book['Sheet']

    name = input("Product name >>> ")
    qty = input("Current quantity >>> ")
    max = Sheet.max_row + 1

    NextID = Book.active.cell(row=1, column=999).value

    Sheet.cell(row=max, column=1).value = NextID
    Sheet.cell(row=max, column=2).value = name
    Sheet.cell(row=max, column=3).value = int(qty)

    Book.active.cell(row=1, column=999).value = Book.active.cell(row=1, column=999).value + 1

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
    ID = input('Input ID of product >>> ')
    i = 0
    for row in Sheet:
        i = i + 1
        if Book.active.cell(row=i, column=1).value == int(ID):
            Sheet.delete_rows(i)
            Book.save('Data.xlsx')

    # Book.save('Data.xlsx')

    ans = input('Remove more? Y/N >>> ').lower()
    if ans == "y":
        product_remove()
    elif ans == "n":
        startScreen()
    else:
        print('Unknown command')
        startScreen()

def clearscreen():
    i = 0
    while i < 20:
        i = i + 1
        print("""
        
        
        
        
        
        """)
