import sys

from PyQt5. QtWidgets import QApplication, QWidget

class Exam(QWidget):

    def __init__(self):
        super().__init__()
        self.initui()
    def initui(self):
        self.show()

app = QApplication(sys.argv)
win = Exam()
sys.exit()

