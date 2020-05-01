from openpyxl import *
from PyQt5 import QtCore, QtGui, QtWidgets
import win32com.client as win

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(381, 185)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.maas = QtWidgets.QLineEdit(self.centralwidget)
        self.maas.setGeometry(QtCore.QRect(20, 70, 91, 33))
        self.maas.setObjectName("maas")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(130, 50, 67, 19))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(220, 50, 67, 19))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(300, 50, 67, 19))
        self.label_3.setObjectName("label_3")
        self.ocak = QtWidgets.QLineEdit(self.centralwidget)
        self.ocak.setGeometry(QtCore.QRect(120, 70, 61, 33))
        self.ocak.setObjectName("ocak")
        self.subat = QtWidgets.QLineEdit(self.centralwidget)
        self.subat.setGeometry(QtCore.QRect(210, 70, 61, 33))
        self.subat.setObjectName("subat")
        self.mart = QtWidgets.QLineEdit(self.centralwidget)
        self.mart.setGeometry(QtCore.QRect(280, 70, 61, 33))
        self.mart.setObjectName("mart")
        self.hesapla = QtWidgets.QPushButton(self.centralwidget)
        self.hesapla.setGeometry(QtCore.QRect(120, 10, 100, 31))
        self.hesapla.setObjectName("hesapla")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 381, 25))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.hesapla.clicked.connect(self.Hesapla)
    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Persona Non Grata"))
        self.label.setText(_translate("MainWindow", "Ocak"))
        self.label_2.setText(_translate("MainWindow", "Şubat"))
        self.label_3.setText(_translate("MainWindow", "Mart"))
        self.hesapla.setText(_translate("MainWindow", "Hesapla"))

    def Hesapla(self):
        self.kitap = load_workbook("deneme.xlsx")
        self.sayfa1 = self.kitap.active
        self.sayfa1["C5"] = int(self.maas.text())
        self.kitap.save("deneme.xlsx")
        self.kitap.close()

        xl = win.DispatchEx("Excel.Application")
        wb = xl.workbooks.open("C:\\Users\\Yüksel\\Desktop\\Yeni klasör\\deneme.xlsx")

        xl.Visible = True
        wb.Save()
        wb.Close()
        xl.Quit()

        self.kitap = load_workbook("deneme.xlsx", data_only=True)
        self.sayfa1 = self.kitap.active

        self.a = str(self.sayfa1["D5"].value)
        self.ocak.setText(self.a)
        self.subat.setText(str(self.sayfa1["E5"].value))

        self.kitap.close()
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())