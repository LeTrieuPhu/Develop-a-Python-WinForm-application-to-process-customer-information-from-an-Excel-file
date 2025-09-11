# open_selected_file()  | mở xem file đã chọn
# open_file()           | mở file
# Xoa_File()            | xóa tên file khỏi list, không xóa file trong hệ thống
# is_file_locked()      | kiểm tra file có đang mở không
# is_file_processed     | kiểm tra file đã được xử lý chưa

# Xu_Ly_File            | xử lý file được chọn
# Cap_Nhat_Thong_Tin()  | cập nhật thông tin mới
# Gop_File
import os
import platform
from pathlib import Path
import subprocess
from PyQt6 import QtCore, QtGui, QtWidgets
import XuLyFileKH

color_text_Title = 'color: rgb(255, 170, 0);'
color_text_Nut = 'color: rgb(0, 0, 100);'
cot_1 = 10
cot_2 = 250
cot_3 = 500
cot_4 = 680
khoi_3 = 470
khoi_4 = 470

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(750, 780)
        MainWindow.setAcceptDrops(True)
        MainWindow.setStyleSheet("background-color: rgb(0, 0, 100);")

        self.centralwidget = QtWidgets.QWidget(parent=MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.QuanLy = QtWidgets.QLabel(parent=self.centralwidget)
        self.QuanLy.setGeometry(QtCore.QRect(cot_1, 10, 180, 60))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.QuanLy.setFont(font)
        self.QuanLy.setStyleSheet(f"{color_text_Title}")
        self.QuanLy.setObjectName("QuanLy")
        self.NutChonFileQLCN = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutChonFileQLCN.setGeometry(QtCore.QRect(cot_1, 60, 80, 30))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.NutChonFileQLCN.setFont(font)
        self.NutChonFileQLCN.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutChonFileQLCN.setObjectName("NutChonFileQLCN")
        self.listFileQL = QtWidgets.QListWidget(parent=self.centralwidget)
        self.listFileQL.setGeometry(QtCore.QRect(cot_1, 90, 220, 100))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.listFileQL.setFont(font)
        self.listFileQL.setStyleSheet(f"background-color: rgb(255, 255, 255); {color_text_Nut}")
        self.listFileQL.setAcceptDrops(True)
        self.listFileQL.setObjectName("listFileQL")
        self.NutXemFile = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutXemFile.setGeometry(QtCore.QRect(cot_1, 200, 220, 80))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutXemFile.setFont(font)
        self.NutXemFile.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutXemFile.setObjectName("NutXemFile")
        self.NutXoaFileQLCN = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutXoaFileQLCN.setGeometry(QtCore.QRect(cot_1, 290, 220, 80))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutXoaFileQLCN.setFont(font)
        self.NutXoaFileQLCN.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutXoaFileQLCN.setObjectName("NutXoaFileQLCN")
        self.NutXuLyFile = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutXuLyFile.setGeometry(QtCore.QRect(cot_1, 380, 220, 80))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutXuLyFile.setFont(font)
        self.NutXuLyFile.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutXuLyFile.setObjectName("NutXuLyFile")
        

        self.KhachHang = QtWidgets.QLabel(parent=self.centralwidget)
        self.KhachHang.setGeometry(QtCore.QRect(cot_2, 10, 240, 60))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.KhachHang.setFont(font)
        self.KhachHang.setObjectName("KhachHang")
        self.KhachHang.setStyleSheet(f"{color_text_Title}")
        self.NutChonFileKH = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutChonFileKH.setGeometry(QtCore.QRect(cot_2, 60, 80, 30))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.NutChonFileKH.setFont(font)
        self.NutChonFileKH.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutChonFileKH.setObjectName("NutChonFileKH")
        self.listFileKH = QtWidgets.QListWidget(parent=self.centralwidget)
        self.listFileKH.setGeometry(QtCore.QRect(cot_2, 90, 230, 370))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.listFileKH.setFont(font)
        self.listFileKH.setStyleSheet(f"background-color: rgb(255, 255, 255);  {color_text_Nut}")
        self.listFileKH.setAcceptDrops(True)
        self.listFileKH.setObjectName("listFileKH")

        self.KhachHangCN = QtWidgets.QLabel(parent=self.centralwidget)
        self.KhachHangCN.setGeometry(QtCore.QRect(cot_1, khoi_3, 220, 60))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.KhachHangCN.setFont(font)
        self.KhachHangCN.setStyleSheet(f"{color_text_Title}")
        self.KhachHangCN.setObjectName("KhachHangCN")
        self.NutChonFileKHCN = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutChonFileKHCN.setGeometry(QtCore.QRect(cot_1, khoi_3+50, 80, 30))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.NutChonFileKHCN.setFont(font)
        self.NutChonFileKHCN.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutChonFileKHCN.setObjectName("NutChonFileKHCN")
        self.listFileKHCN = QtWidgets.QListWidget(parent=self.centralwidget)
        self.listFileKHCN.setGeometry(QtCore.QRect(cot_1, khoi_3+80, 220, 100))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.listFileKHCN.setFont(font)
        self.listFileKHCN.setStyleSheet(f"background-color: rgb(255, 255, 255); {color_text_Nut}")
        self.listFileKHCN.setAcceptDrops(True)
        self.listFileKHCN.setObjectName("listFileKHCN")
        self.NutCapNhat = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutCapNhat.setGeometry(QtCore.QRect(cot_1, khoi_3+190, 220, 100))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutCapNhat.setFont(font)
        self.NutCapNhat.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutCapNhat.setObjectName("NutCapNhat")


        self.KhachHangGop = QtWidgets.QLabel(parent=self.centralwidget)
        self.KhachHangGop.setGeometry(QtCore.QRect(cot_2, khoi_4, 230, 60))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.KhachHangGop.setFont(font)
        self.KhachHangGop.setStyleSheet(f"{color_text_Title}")
        self.KhachHangGop.setObjectName("KhachHangGop")
        self.NutChonFileGop = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutChonFileGop.setGeometry(QtCore.QRect(cot_2, khoi_4+50, 80, 30))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.NutChonFileGop.setFont(font)
        self.NutChonFileGop.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutChonFileGop.setObjectName("NutChonFileGop")
        self.listFileGop = QtWidgets.QListWidget(parent=self.centralwidget)
        self.listFileGop.setGeometry(QtCore.QRect(cot_2, khoi_4+80, 230, 100))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.listFileGop.setFont(font)
        self.listFileGop.setStyleSheet(f"background-color: rgb(255, 255, 255); {color_text_Nut}")
        self.listFileGop.setAcceptDrops(True)
        self.listFileGop.setObjectName("listFileGop")
        self.NutGop = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutGop.setGeometry(QtCore.QRect(cot_2, khoi_4+190, 230, 100))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutGop.setFont(font)
        self.NutGop.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutGop.setObjectName("NutGop")

        self.LichSulabel = QtWidgets.QLabel(parent=self.centralwidget)
        self.LichSulabel.setGeometry(QtCore.QRect(cot_3+5, 10, 150, 60))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.LichSulabel.setFont(font)
        self.LichSulabel.setStyleSheet(f"{color_text_Title}")
        self.LichSulabel.setObjectName("LichSulabel")
        self.listFileLS = QtWidgets.QListWidget(parent=self.centralwidget)
        self.listFileLS.setGeometry(QtCore.QRect(cot_3, 90, 240, 370))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.listFileLS.setFont(font)
        self.listFileLS.setStyleSheet(f"background-color: rgb(255, 255, 255); {color_text_Nut}")
        self.listFileLS.setObjectName("listFileLS")

        self.NutBack = QtWidgets.QPushButton(parent=self.centralwidget)
        self.NutBack.setGeometry(QtCore.QRect(cot_3, 610, 220, 100))
        font = QtGui.QFont()
        font.setPointSize(18)
        self.NutBack.setFont(font)
        self.NutBack.setStyleSheet(f"background-color: rgb(210, 255, 255); {color_text_Nut}")
        self.NutBack.setObjectName("NutBack")


        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(parent=MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 629, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(parent=MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Xử Lý Mail Khách Hàng"))
        self.QuanLy.setText(_translate("MainWindow", "Tệp Quản Lý"))
        self.NutChonFileQLCN.setText(_translate("MainWindow", "Chọn File"))
        self.KhachHang.setText(_translate("MainWindow", "Tệp Khách Hàng"))
        self.NutChonFileKH.setText(_translate("MainWindow", "Chọn File"))
        self.NutXoaFileQLCN.setText(_translate("MainWindow", "Xóa File"))
        self.NutXuLyFile.setText(_translate("MainWindow", "Xử Lý File"))
        self.NutXemFile.setText(_translate("MainWindow", "Xem File"))
        self.NutBack.setText(_translate("MainWindow", "Trở Về"))

        self.KhachHangCN.setText(_translate("MainWindow", "Cập Nhật Thông Tin"))
        self.NutChonFileKHCN.setText(_translate("MainWindow", "Chọn File"))
        self.NutCapNhat.setText(_translate("MainWindow", "Cập Nhật"))
        
        self.KhachHangGop.setText(_translate("MainWindow", "Gộp Tệp Khách Hàng"))
        self.NutChonFileGop.setText(_translate("MainWindow", "Chọn File"))
        self.NutGop.setText(_translate("MainWindow", "Gộp"))

        self.LichSulabel.setText(_translate("MainWindow", "Lịch Sử Xử Lý"))

# ---------- Thêm MainWindow xử lý drag/drop cho 2 widget ----------
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.changed_cells = []
        self.fw = None

        # cài event filter cho hai QListWidget
        self.ui.listFileKH.installEventFilter(self)
        self.ui.listFileQL.installEventFilter(self)
        self.ui.listFileKHCN.installEventFilter(self)
        self.ui.listFileGop.installEventFilter(self)
        self.ui.listFileLS.installEventFilter(self)

        # đảm bảo widget hiển thị dấu chấp nhận khi di chuột (khi cần)
        self.ui.listFileKH.setAcceptDrops(True)
        self.ui.listFileQL.setAcceptDrops(True)
        self.ui.listFileKHCN.setAcceptDrops(True)
        self.ui.listFileGop.setAcceptDrops(True)
        self.ui.listFileLS.setAcceptDrops(True)

        # trong __init__ của MainWindow sau khi setupUi xong
        self.ui.listFileKH.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.ui.listFileQL.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.ui.listFileKHCN.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.ui.listFileGop.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.ui.listFileLS.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)

        # Xuống dòng tự động
        self.ui.listFileQL.setWordWrap(True)
        self.ui.listFileKH.setWordWrap(True)
        self.ui.listFileKHCN.setWordWrap(True)
        self.ui.listFileGop.setWordWrap(True)
        self.ui.listFileLS.setWordWrap(True)

        # Kết nối sự kiện thay đổi chọn
        self.ui.listFileKH.itemSelectionChanged.connect(self._clear_other_selection_from_KH)
        self.ui.listFileQL.itemSelectionChanged.connect(self._clear_other_selection_from_QL)
        self.ui.listFileKHCN.itemSelectionChanged.connect(self._clear_other_selection_from_KHCN)
        self.ui.listFileGop.itemSelectionChanged.connect(self._clear_other_selection_from_Gop)
        self.ui.listFileLS.itemSelectionChanged.connect(self._clear_other_selection_from_LS)

    # 5 hàm _clear_other_selection thực hiện nhiệm vụ điều hướng con trỏ chọn giữa các danh sách
    # Chọn item của danh sách này thì sẽ bỏ chọn ở tất cả danh sách còn lại
    def _clear_other_selection_from_KH(self):
        """Khi chọn ở listFileKH thì bỏ chọn listFileQL"""
        if self.ui.listFileKH.selectedItems():
            self.ui.listFileQL.clearSelection()
            self.ui.listFileKHCN.clearSelection()
            self.ui.listFileGop.clearSelection()
            self.ui.listFileLS.clearSelection()
            self.fw = self.ui.listFileKH

    def _clear_other_selection_from_QL(self):
        """Khi chọn ở listFileQL thì bỏ chọn listFileKH"""
        if self.ui.listFileQL.selectedItems():
            self.ui.listFileKH.clearSelection()
            self.ui.listFileKHCN.clearSelection()
            self.ui.listFileGop.clearSelection()
            self.ui.listFileLS.clearSelection()
            self.fw = self.ui.listFileQL
    
    def _clear_other_selection_from_KHCN(self):
        """Khi chọn ở listFileKH thì bỏ chọn listFileQL"""
        if self.ui.listFileKHCN.selectedItems():
            self.ui.listFileQL.clearSelection()
            self.ui.listFileKH.clearSelection()
            self.ui.listFileGop.clearSelection()
            self.ui.listFileLS.clearSelection()
            self.fw = self.ui.listFileKHCN

    def _clear_other_selection_from_Gop(self):
        """Khi chọn ở listFileKH thì bỏ chọn listFileQL"""
        if self.ui.listFileGop.selectedItems():
            self.ui.listFileKH.clearSelection()
            self.ui.listFileQL.clearSelection()
            self.ui.listFileKHCN.clearSelection()
            self.ui.listFileLS.clearSelection()
            self.fw = self.ui.listFileGop

    def _clear_other_selection_from_LS(self):
        """Khi chọn ở listFileKH thì bỏ chọn listFileQL"""
        if self.ui.listFileLS.selectedItems():
            self.ui.listFileKH.clearSelection()
            self.ui.listFileQL.clearSelection()
            self.ui.listFileKHCN.clearSelection()
            self.ui.listFileGop.clearSelection()
            self.fw = self.ui.listFileLS

    # Hàm có chức năng xử lý việc kéo thả file vào Widget
    def eventFilter(self, obj, event):
        # drag enter: chấp nhận nếu có file/url
        if event.type() == QtCore.QEvent.Type.DragEnter:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True

        # drag move: cần thiết để có hiệu ứng hover khi kéo
        if event.type() == QtCore.QEvent.Type.DragMove:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True

        # drop: lấy file và thêm vào list tương ứng
        if event.type() == QtCore.QEvent.Type.Drop:
            if event.mimeData().hasUrls():
                urls = event.mimeData().urls()
                # chuyển các QUrl -> đường dẫn hệ thống
                paths = [u.toLocalFile() for u in urls if u.isLocalFile()]
                if paths:
                    # nếu drop vào widget Khách Hàng
                    if obj is self.ui.listFileKH:
                        for p in paths:
                            name = os.path.basename(p)  # chỉ lấy tên file
                            item = QtWidgets.QListWidgetItem(name)
                            # lưu full path vào UserRole để dùng sau
                            item.setData(QtCore.Qt.ItemDataRole.UserRole, p)
                            item.setToolTip(p)  # hiển thị full path khi hover
                            self.ui.listFileKH.addItem(item)
                    # nếu drop vào widget Quản Lý
                    elif obj is self.ui.listFileQL:
                        for p in paths:
                            name = os.path.basename(p)
                            item = QtWidgets.QListWidgetItem(name)
                            item.setData(QtCore.Qt.ItemDataRole.UserRole, p)
                            item.setToolTip(p)
                            self.ui.listFileQL.addItem(item)
                    # nếu drop vào widget Cập Nhật
                    elif obj is self.ui.listFileKHCN:
                        for p in paths:
                            name = os.path.basename(p)
                            item = QtWidgets.QListWidgetItem(name)
                            item.setData(QtCore.Qt.ItemDataRole.UserRole, p)
                            item.setToolTip(p)
                            self.ui.listFileKHCN.addItem(item)
                    # nếu drop vào widget Gộp File
                    elif obj is self.ui.listFileGop:
                        for p in paths:
                            name = os.path.basename(p)
                            item = QtWidgets.QListWidgetItem(name)
                            item.setData(QtCore.Qt.ItemDataRole.UserRole, p)
                            item.setToolTip(p)
                            self.ui.listFileGop.addItem(item)
                event.acceptProposedAction()
                return True

        # không bắt được sự kiện, trả về hành vi mặc định
        return super().eventFilter(obj, event)

    # Chỉ ra item cần mở file và lấy full đường dẫn của file
    def open_selected_file(self):
        """Mở file của item đang chọn.
        Quy tắc:
         - Nếu widget có focus và có item chọn -> mở item đó
         - Nếu không, ưu tiên kiểm tra listFileKH, rồi listFileQL
        """
        # fw là biến dùng để chỉ ra là con trỏ chuột đang chọn vào widget nào, self.fw đã được xử lý ở các hàm _clear_other_selection
        fw = self.fw
        candidate = None

        # Lấy ra item tương ứng với Widget được trỏ tới
        if fw is self.ui.listFileKH:
            candidate = self.ui.listFileKH.currentItem()
        elif fw is self.ui.listFileQL:
            candidate = self.ui.listFileQL.currentItem()
        elif fw is self.ui.listFileKHCN:
            candidate = self.ui.listFileKHCN.currentItem()
        elif fw is self.ui.listFileGop:
            candidate = self.ui.listFileGop.currentItem()
        elif fw is self.ui.listFileLS:
            candidate = self.ui.listFileLS.currentItem()

        if not candidate:
            self.show_information('Chưa chọn file để mở')
            return

        full = candidate.data(QtCore.Qt.ItemDataRole.UserRole)
        if not full:
            self.show_warning('Không tìm được đường dẫn đầy đủ của file')
            return

        self.open_file(full)

    # Hàm mở file chính
    def open_file(self, path):
        """Mở file bằng ứng dụng mặc định hệ điều hành.
        Hỗ trợ Windows, macOS, Linux (xdg-open).
        """
        if not os.path.exists(path):
            self.show_warning(f'File không tồn tại:\n{path}')
            return
        try:
            system = platform.system()
            if system == "Windows":
                os.startfile(path)  # type: ignore[attr-defined]
            elif system == "Darwin":  # macOS
                subprocess.run(["open", path], check=False)
            else:  # Giả sử Linux/Unix
                subprocess.run(["xdg-open", path], check=False)
        except Exception as e:
            self.show_warning(f'Không thể mở file:\n{path}\n\n{e}')

    # Hàm có nhiệm vụ xóa item được chọn khỏi danh sách tương ứng
    def Xoa_File(self):
        fw = self.fw

        if fw is self.ui.listFileKH:
            listItems = self.ui.listFileKH.selectedItems()
            if not listItems:  # không có item nào được chọn
                return
            for item in listItems:
                self.ui.listFileKH.takeItem(self.ui.listFileKH.row(item))
        elif fw is self.ui.listFileQL:
            listItems = self.ui.listFileQL.selectedItems()
            if not listItems:  # không có item nào được chọn
                return
            for item in listItems:
                self.ui.listFileQL.takeItem(self.ui.listFileQL.row(item))
        elif fw is self.ui.listFileKHCN:
            listItems = self.ui.listFileKHCN.selectedItems()
            if not listItems:  # không có item nào được chọn
                return
            for item in listItems:
                self.ui.listFileKHCN.takeItem(self.ui.listFileKHCN.row(item))
        elif fw is self.ui.listFileGop:
            listItems = self.ui.listFileGop.selectedItems()
            if not listItems:  # không có item nào được chọn
                return
            for item in listItems:
                self.ui.listFileGop.takeItem(self.ui.listFileGop.row(item))
        elif fw is self.ui.listFileLS:
            self.show_warning('Không thể xóa lịch sử')
    
    # 4 hàm chọn file, thực hiện nhiệm vụ chọn file từ hệ thông lên widget tương ứng
    def chon_file_tu_o_dia_cho_QLCN(self):
        if self.ui.listFileQL.count() > 0:
            self.show_warning('Chỉ được chọn một file quản lý\nVui lòng xóa file cũ trước khi thêm file mới')
            return
        
        """Mở hộp thoại chọn file và thêm vào listFileQL (hoặc listFileKH)"""
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn file", "", "All Files (*.*)"
        )
        if file_path:
            name = os.path.basename(file_path)
            item = QtWidgets.QListWidgetItem(name)
            item.setData(QtCore.Qt.ItemDataRole.UserRole, file_path)
            item.setToolTip(file_path)

            # 👉 ở đây bạn chọn muốn đưa file vào list nào
            self.ui.listFileQL.addItem(item)

    def chon_file_tu_o_dia_cho_KH(self):
        """Mở hộp thoại chọn file và thêm vào listFileKH (không trùng lặp)"""
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn file", "", "All Files (*.*)"
        )
        
        if file_path:
            name = os.path.basename(file_path)

            # 🔍 kiểm tra xem đã có file trong list chưa
            for i in range(self.ui.listFileKH.count()):
                item = self.ui.listFileKH.item(i)
                if item.data(QtCore.Qt.ItemDataRole.UserRole) == file_path:
                    self.show_information(f'File {name} đã tồn tại trong danh sách!')
                    return

            new_item = QtWidgets.QListWidgetItem(name)
            new_item.setData(QtCore.Qt.ItemDataRole.UserRole, file_path)
            new_item.setToolTip(file_path)
            self.ui.listFileKH.addItem(new_item)
    
    def chon_file_tu_o_dia_cho_KHCN(self):
        """Mở hộp thoại chọn file và thêm vào listFileKHCN (không trùng lặp)"""
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn file", "", "All Files (*.*)"
        )
        
        if file_path:
            name = os.path.basename(file_path)

            # 🔍 kiểm tra xem đã có file trong list chưa
            for i in range(self.ui.listFileKHCN.count()):
                item = self.ui.listFileKHCN.item(i)
                if item.data(QtCore.Qt.ItemDataRole.UserRole) == file_path: 
                    self.show_information(f'File {name} đã tồn tại trong danh sách!') 
                    return

            new_item = QtWidgets.QListWidgetItem(name)
            new_item.setData(QtCore.Qt.ItemDataRole.UserRole, file_path)
            new_item.setToolTip(file_path)
            self.ui.listFileKHCN.addItem(new_item)

    def chon_file_tu_o_dia_cho_Gop(self):
        """Mở hộp thoại chọn file và thêm vào listFileKH (không trùng lặp)"""
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Chọn file", "", "All Files (*.*)"
        )
        
        if file_path:
            name = os.path.basename(file_path)

            # 🔍 kiểm tra xem đã có file trong list chưa
            for i in range(self.ui.listFileGop.count()):
                item = self.ui.listFileGop.item(i)
                if item.data(QtCore.Qt.ItemDataRole.UserRole) == file_path:  
                    self.show_information(f'File {name} đã tồn tại trong danh sách!')
                    return

            new_item = QtWidgets.QListWidgetItem(name)
            new_item.setData(QtCore.Qt.ItemDataRole.UserRole, file_path)
            new_item.setToolTip(file_path)
            self.ui.listFileGop.addItem(new_item)

    def is_file_locked(self, filepath):
        """Kiểm tra file có đang bị khóa (ví dụ đang mở trong Excel) không"""
        if not os.path.exists(filepath):
            return False
        try:
            # thử mở để ghi (exclusive)
            with open(filepath, "a"):
                return False
        except PermissionError:
            return True
        
    def is_file_processed(self, input_path):
        input_stem = Path(str(input_path)).stem.lower()   # lấy tên không có .ext, chuyển về lowercase
        for i in range(self.ui.listFileLS.count()):
            item_stem = Path(str(self.ui.listFileLS.item(i))).stem.lower()
            if input_stem == item_stem:
                return False
        return True

    def show_warning(self, text):
        msg = QtWidgets.QMessageBox(self)
        msg.setWindowTitle("Cảnh báo")

        # Text bạn tự định dạng bằng HTML hoặc plain text
        msg.setText(f'{text}')
        msg.setIcon(QtWidgets.QMessageBox.Icon.Warning)

        # Chỉ đổi nền toàn bộ hộp thoại sang trắng
        msg.setStyleSheet("""
            QMessageBox {
                background-color: rgb(255, 255, 255);
            }
            QLabel {
                background-color: white;
            }
            QPushButton {
                background-color: white;
            }
        """)

        msg.exec()

    def show_information(self, text):
        msg = QtWidgets.QMessageBox(self)
        msg.setWindowTitle("Thông báo")

        # Text bạn tự định dạng bằng HTML hoặc plain text
        msg.setText(f'{text}')
        msg.setIcon(QtWidgets.QMessageBox.Icon.Information)

        # Chỉ đổi nền toàn bộ hộp thoại sang trắng
        msg.setStyleSheet("""
            QMessageBox {
                background-color: rgb(255, 255, 255);
            }
                          QLabel {
                background-color: white;
            }
            QPushButton {
                background-color: white;
            }
        """)

        msg.exec()

    def Xu_Ly_File(self):
        # đường dẫn input file
        input_path = self.ui.listFileKH.currentItem()
        # đường dẫn output file
        output_path = self.ui.listFileQL.item(0)

        if not input_path:
            self.show_warning('Vui lòng chọn File cần xử lý trước khi xử lý')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", '<span style="color: rgb(255, 170, 0);">Vui lòng chọn File cần xử lý trước khi xử lý</span>')
            return
        else: 
            input_path = input_path.data(QtCore.Qt.ItemDataRole.UserRole)
    
        if not output_path:
            self.show_warning('Vui lòng chọn File quản lý trước khi xử lý')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", '<span style="color: rgb(255, 170, 0);">Vui lòng chọn File quản lý trước khi xử lý</span>')
            return
        else:
            output_path = output_path.data(QtCore.Qt.ItemDataRole.UserRole)

        if not self.is_file_processed(input_path):
            self.show_information(f'File {Path(str(os.path.basename(input_path))).stem.lower()} đã được xử lý')
            # QtWidgets.QMessageBox.information(self, "Thông báo", f'<span style="color: rgb(255, 170, 0);">File {os.path.basename(output_path)} đã được xử lý</span>')
            return
        
        # 🔒 kiểm tra file output có đang mở không
        if self.is_file_locked(output_path):
            self.show_warning(f'File {os.path.basename(output_path)} đang mở. Vui lòng đóng file trước khi xử lý')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", f'<span style="color: rgb(255, 170, 0);">File {os.path.basename(output_path)} đang mở. Vui lòng đóng file trước khi xử lý</span>')
            return
        if XuLyFileKH.Loc_Thong_Tin(input_path, output_path):
            name = os.path.basename(input_path)
            item = QtWidgets.QListWidgetItem(name)
            item.setData(QtCore.Qt.ItemDataRole.UserRole, input_path)
            item.setToolTip(input_path)

            # 👉 ở đây bạn chọn muốn đưa file vào list nào
            self.ui.listFileLS.addItem(item)
            self.Xoa_File()
            self.show_information('Mã đơn hàng đã được thêm')
            # QtWidgets.QMessageBox.information(self, "Thông báo", '<span style="color: rgb(255, 170, 0);">Mã đơn hàng đã được thêm</span>')
        else:
            self.show_information('Mã đơn hàng đã tồn tại')
            # QtWidgets.QMessageBox.information(self, "Thông báo", '<span style="color: rgb(255, 170, 0);">Mã đơn hàng đã tồn tại</span>')

    def Cap_Nhat_Thong_Tin(self):
        # đường dẫn cập nhật file
        CapNhat_path = self.ui.listFileKHCN.currentItem()
        # đường dẫn output file
        output_path = self.ui.listFileQL.item(0)

        if not CapNhat_path:
            self.show_warning('Vui lòng chọn File Cần cập nhật trước khi cập nhật')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", '<span style="color: rgb(255, 170, 0);">Vui lòng chọn File Cần cập nhật trước khi cập nhật</span>')
            return
        else: 
            CapNhat_path = CapNhat_path.data(QtCore.Qt.ItemDataRole.UserRole)

        if not output_path:
            self.show_warning('Vui lòng chọn File quản lý trước khi cập nhật')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", '<span style="color: rgb(255, 170, 0);">Vui lòng chọn File quản lý trước khi cập nhật</span>')
            return
        else:
            output_path = output_path.data(QtCore.Qt.ItemDataRole.UserRole)
        
        # 🔒 kiểm tra file output có đang mở không
        if self.is_file_locked(output_path):
            self.show_warning(f'File {os.path.basename(output_path)} đang mở. Vui lòng đóng file trước khi xử lý')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", f'<span style="color: rgb(255, 170, 0);">File {os.path.basename(output_path)} đang mở. Vui lòng đóng file trước khi xử lý</span>')
            return
        
        if XuLyFileKH.Cap_Nhat_Thong_Tin(CapNhat_path, output_path, self.changed_cells):
            name = os.path.basename(CapNhat_path)
            item = QtWidgets.QListWidgetItem(name)
            item.setData(QtCore.Qt.ItemDataRole.UserRole, CapNhat_path)
            item.setToolTip(CapNhat_path)

            # 👉 ở đây bạn chọn muốn đưa file vào list nào
            self.ui.listFileLS.addItem(item)
            self.Xoa_File()
            self.show_information('Dữ liệu của mã đơn hàng đã được cập nhật')
            # QtWidgets.QMessageBox.information(self, "Thông báo", '<span style="color: rgb(255, 170, 0);">Dữ liệu của mã đơn hàng đã được cập nhật</span>')
        else:
            self.show_information('Mã đơn hàng không tồn tại')
            # QtWidgets.QMessageBox.information(self, "Thông báo", '<span style="color: rgb(255, 170, 0);">Mã đơn hàng chưa được cập nhật</span>')

    def Gop_File(self):
        # đường dẫn cập nhật file
        list_path_file = []
        for i in range(self.ui.listFileGop.count()):  
            list_path_file.append(self.ui.listFileGop.item(i).data(QtCore.Qt.ItemDataRole.UserRole))

        # đường dẫn output file
        output_path = self.ui.listFileQL.item(0)

        if not output_path:
            self.show_warning('Vui lòng chọn File quản lý trước khi cập gộp')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", '<span style="color: rgb(255, 170, 0);">Vui lòng chọn File quản lý trước khi cập nhật</span>')
            return
        else:
            output_path = output_path.data(QtCore.Qt.ItemDataRole.UserRole)

        # 🔒 kiểm tra file output có đang mở không
        if self.is_file_locked(output_path):
            self.show_warning(f'File {os.path.basename(output_path)} đang mở. Vui lòng đóng file trước khi xử lý')
            # QtWidgets.QMessageBox.warning(self, "Cảnh báo", f'<span style="color: rgb(255, 170, 0);">File {os.path.basename(output_path)} đang mở. Vui lòng đóng file trước khi xử lý</span>')
            return
        
        if XuLyFileKH.Gop_File(list_path_file, output_path):
            for input_path in list_path_file:
                name = os.path.basename(input_path)
                item = QtWidgets.QListWidgetItem(name)
                item.setData(QtCore.Qt.ItemDataRole.UserRole, input_path)
                item.setToolTip(input_path)
                # 👉 ở đây bạn chọn muốn đưa file vào list nào
                self.ui.listFileLS.addItem(item)
            self.ui.listFileGop.clear()
            self.show_information('Dữ liệu của mã đơn hàng đã được gộp và thêm')
            # QtWidgets.QMessageBox.information(self, "Thông báo", '<span style="color: rgb(255, 170, 0);">Dữ liệu của mã đơn hàng đã được gộp và thêm</span>')
        else:
            self.show_information('Mã đơn hàng chưa được gộp và thêm')
            # QtWidgets.QMessageBox.information(self, "Thông báo", '<span style="color: rgb(255, 170, 0);">Mã đơn hàng chưa được gộp và thêm')
if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())
