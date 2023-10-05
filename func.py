import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import PySimpleGUI as sg
from func import startScreen
import os

# for func.py ---------------------
from main import Book,Sheet



def startScreen():
    Choice = input('Choice >>> ').lower()

    if Choice == "add":
        print('add')
    elif Choice == "read":
        print('add')
    elif Choice == "write":
        print('add')
    elif Choice == "help":
        print('add')
    elif Choice == "exit":
        print('exiting...')
    else:
        print('unknown command, try again')
        startScreen()