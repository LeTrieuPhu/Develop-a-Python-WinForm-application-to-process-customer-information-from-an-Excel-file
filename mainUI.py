import os, sys
from PyQt6 import QtCore, QtGui, QtWidgets
import HomeUI, XuLyFileKHUI

# các biến list dùng để lưu lại danh sách item trong widget trước khi quay về trang home
listFileKH = None
listFileQLCN = None
listFileCN = None
listFileGop = None

# biến ui là biến giao diện chính của ứng dụng
ui = ''
app = QtWidgets.QApplication(sys.argv)
MainWindow = QtWidgets.QMainWindow()

# hàm khởi chạy giao diện chính của ứng dụng
def home():
    global ui
    ui = HomeUI.Ui_MainWindow()
    ui.setupUi(MainWindow)

    # xử lý nút chuyển hướng sang trang xử lý file
    ui.NutXuLyFileKH.clicked.connect(Xu_Ly_File_KH)

    MainWindow.show()

def Xu_Ly_File_KH():
    global ui, listFileKH, listFileQLCN, listFileCN, listFileGop
    ui = XuLyFileKHUI.MainWindow()

    # Khi vào trang xử lý, kiểm tra các list có dữ liệu không, nếu có thì thêm tất cả vào widget
    if listFileKH is not None:
        for item in listFileKH:
            ui.ui.listFileKH.addItem(item)
    if listFileQLCN is not None:
        for item in listFileQLCN:
            ui.ui.listFileQL.addItem(item)
    if listFileCN is not None:
        for item in listFileCN:
            ui.ui.listFileKHCN.addItem(item)
    if listFileGop is not None:
        for item in listFileGop:
            ui.ui.listFileGop.addItem(item)

    # Nút chọn và thêm file từ hệ thống vào Widget
    ui.ui.NutChonFileQLCN.clicked.connect(ui.chon_file_tu_o_dia_cho_QLCN)       
    ui.ui.NutChonFileKH.clicked.connect(ui.chon_file_tu_o_dia_cho_KH) 
    ui.ui.NutChonFileKHCN.clicked.connect(ui.chon_file_tu_o_dia_cho_KHCN)
    ui.ui.NutChonFileGop.clicked.connect(ui.chon_file_tu_o_dia_cho_Gop)

    # Nút xử lý chức năng xem, xóa, xử lý, trở về home, cập nhật và gộp file (chỉ xóa item khỏi list, không xóa file gốc)
    ui.ui.NutXemFile.clicked.connect(ui.open_selected_file)
    ui.ui.NutXoaFileQLCN.clicked.connect(ui.Xoa_File)
    ui.ui.NutXuLyFile.clicked.connect(ui.Xu_Ly_File)
    ui.ui.NutBack.clicked.connect(Back_to_home)
    ui.ui.NutCapNhat.clicked.connect(ui.Cap_Nhat_Thong_Tin)
    ui.ui.NutGop.clicked.connect(ui.Gop_File)
    ui.show()

# hàm Back thực hiện lưu lại các item trong Widget vào các biến để khôi phục sau này
def Back_to_home():
    global listFileKH, listFileQLCN, listFileCN, listFileGop
    # Lưu lại danh sách email (copy item text ra)
    listFileKH = [ui.ui.listFileKH.item(i).text() for i in range(ui.ui.listFileKH.count())]
    listFileQLCN = [ui.ui.listFileQL.item(i).text() for i in range(ui.ui.listFileQL.count())]
    listFileCN = [ui.ui.listFileKHCN.item(i).text() for i in range(ui.ui.listFileKHCN.count())]
    listFileGop = [ui.ui.listFileGop.item(i).text() for i in range(ui.ui.listFileGop.count())]
    # Quay về giao diện home
    home()
# Run app
home()
sys.exit(app.exec())

