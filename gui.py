from PyQt6.QtWidgets import QMainWindow, QApplication, QPushButton

import sys


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("EasyBooks")

        button = QPushButton("Button")
        button.pressed.connect(self.close)

        self.setCentralWidget(button)
        self.show()


app = QApplication(sys.argv)
w = MainWindow()
app.exec()