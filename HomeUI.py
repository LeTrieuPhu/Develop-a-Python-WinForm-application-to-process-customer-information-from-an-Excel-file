# Hiển thị giao diện chính của ứng dụng
# Hiện tại chỉ có một tính năng xử lý đầu vào từ khách hàng

from PyQt6 import QtCore, QtGui, QtWidgets

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(286, 279)
        MainWindow.setStyleSheet("background-color: rgb(0, 0, 100);")

        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.NutXuLyFileKH = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutXuLyFileKH.setGeometry(QtCore.QRect(10, 10, 260, 100))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutXuLyFileKH.setFont(font)
        self.NutXuLyFileKH.setStyleSheet("background-color: rgb(210, 255, 255); color: rgb(0, 0, 100);")
        self.NutXuLyFileKH.setObjectName("NutXuLyFileKH")
        self.NutXemLichSu = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutXemLichSu.setGeometry(QtCore.QRect(10, 120, 260, 100))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutXemLichSu.setFont(font)
        self.NutXemLichSu.setStyleSheet("background-color: rgb(210, 255, 255); color: rgb(0, 0, 100);")
        self.NutXemLichSu.setObjectName("NutXemLichSu")

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 286, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Home"))
        self.NutXuLyFileKH.setText(_translate("MainWindow", "Xử Lý File\n"
"Khách Hàng"))
        self.NutXemLichSu.setText(_translate("MainWindow", "Lịch Sử Xử Lý"))


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec())
