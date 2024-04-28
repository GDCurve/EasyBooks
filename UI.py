# Sis ir ar qt designer taisiits

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QTableWidget, QTableWidgetItem, QDialog, QVBoxLayout, QLabel, QLineEdit, QPushButton
from PyQt5.QtCore import Qt
from func import load, finder, save
from PyQt5.QtCore import QCoreApplication

class IzdzestDialogs(QDialog):
    def __init__(self, parent=None):
        super(IzdzestDialogs, self).__init__(parent)

        self.setWindowTitle("Remove Product")

        self.id_label = QLabel("Enter Product ID:")
        self.id_input = QLineEdit()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")

        layout = QVBoxLayout()
        layout.addWidget(self.id_label)
        layout.addWidget(self.id_input)
        layout.addWidget(self.ok_button)
        layout.addWidget(self.cancel_button)
        self.setLayout(layout)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

class PievienotDialogs(QDialog):
    def __init__(self, parent=None):
        super(PievienotDialogs, self).__init__(parent)

        self.setWindowTitle("Add Product")

        self.id_label = QLabel("Enter Product Count:")
        self.id_input = QLineEdit()
        self.name_label = QLabel("Enter Product Name:")
        self.name_input = QLineEdit()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")

        layout = QVBoxLayout()
        layout.addWidget(self.id_label)
        layout.addWidget(self.id_input)
        layout.addWidget(self.name_label)
        layout.addWidget(self.name_input)
        layout.addWidget(self.ok_button)
        layout.addWidget(self.cancel_button)
        self.setLayout(layout)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

class EditDialogs(QDialog):
    def __init__(self, parent=None):
        super(EditDialogs, self).__init__(parent)

        self.setWindowTitle("Edit Product")

        self.id_label = QLabel("Enter Product ID:")
        self.id_input = QLineEdit()
        self.count_label = QLabel("Count By How Much?:")
        self.count_input = QLineEdit()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")

        layout = QVBoxLayout()
        layout.addWidget(self.id_label)
        layout.addWidget(self.id_input)
        layout.addWidget(self.count_label)
        layout.addWidget(self.count_input)
        layout.addWidget(self.ok_button)
        layout.addWidget(self.cancel_button)
        self.setLayout(layout)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(702, 576)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")


        self.table = QTableWidget(self.centralwidget)
        self.table.setGeometry(QtCore.QRect(10, 10, 451, 521))
        self.table.setObjectName("table")

        self.update_table()

        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setGeometry(QtCore.QRect(480, 10, 201, 51))
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.atvert_izdzest)

        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.setGeometry(QtCore.QRect(480, 70, 201, 51))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_2.clicked.connect(self.atvert_pievienot)

        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.setGeometry(QtCore.QRect(480, 130, 201, 51))
        self.pushButton_3.setObjectName("pushButton_3")
        self.pushButton_3.clicked.connect(self.atvert_edit)

        self.exitButton = QtWidgets.QPushButton(self.centralwidget)
        self.exitButton.setGeometry(QtCore.QRect(480, 480, 201, 51))
        self.exitButton.setObjectName("exitButton")
        self.exitButton.setText("Exit")
        self.exitButton.clicked.connect(QCoreApplication.instance().quit)

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 702, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def update_table(self):
        Book, Sheet = load()
        self.table.setRowCount(Sheet.max_row)
        self.table.setColumnCount(3)
        #platums

        self.table.setColumnWidth(0, 30)
        self.table.setColumnWidth(1, 300)
        self.table.setColumnWidth(2, 70)

        for i, row in enumerate(Sheet.iter_rows(values_only=True), start=1):
            for j, value in enumerate(row, start=1):
                item = QTableWidgetItem(str(value))
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.table.setItem(i - 1, j - 1, item)
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "EasyBooks"))
        self.pushButton.setText(_translate("MainWindow", "Remove Product"))
        self.pushButton_2.setText(_translate("MainWindow", "Add Product"))
        self.pushButton_3.setText(_translate("MainWindow", "Edit Count"))

    def atvert_izdzest(self):
        dialog = IzdzestDialogs(self.centralwidget)
        result = dialog.exec_()

        if result == QDialog.Accepted:
            id = dialog.id_input.text()
            Book, Sheet = load()

            name, qty, i = finder(id)
            if id.isdigit() == True:
                Sheet.delete_rows(i)
                Book.save('Data.xlsx')
                self.update_table()
            else:
                self.atvert_izdzest()


    def atvert_pievienot(self):
        dialog = PievienotDialogs(self.centralwidget)
        result = dialog.exec_()

        if result == QDialog.Accepted:
            qty = dialog.id_input.text()
            name = dialog.name_input.text()
            Book, Sheet = load()
            if qty.isdigit() == True:

                max = Sheet.max_row + 1

                NextID = Book.active.cell(row=1, column=999).value

                Sheet.cell(row=max, column=1).value = NextID
                Sheet.cell(row=max, column=2).value = name
                Sheet.cell(row=max, column=3).value = int(qty)

                Book.active.cell(row=1, column=999).value = Book.active.cell(row=1, column=999).value + 1
                Book.save('Data.xlsx')
                self.update_table()
            else:
                self.atvert_pievienot()


    def atvert_edit(self):
        dialog = EditDialogs(self.centralwidget)
        result = dialog.exec_()

        if result == QDialog.Accepted:
            id = dialog.id_input.text()
            count = dialog.count_input.text()
            if id.isdigit() == True or count.isdigit() == True:
                finder(id)
                Book, Sheet = load()
                name, qty, i = finder(id)
                Sheet.cell(row=i, column=3).value = Sheet.cell(row=i, column=3).value + int(count)
                Book.save('Data.xlsx')
                self.update_table()
            else:
                self.atvert_edit()



if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())