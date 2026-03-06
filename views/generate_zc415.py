from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QComboBox, QCheckBox, QAbstractItemView, QTableWidget, \
    QHeaderView, QPushButton, QDesktopWidget, QLabel, QLineEdit, QScrollArea, QWidget, QSizePolicy, QMessageBox, \
    QFileDialog

from http_client import download_zc415_file


class GenerateZC415(QDialog):
    # def __init__(self, username, token):
    def __init__(self):
        super().__init__()
        # self.username = username
        # self.token = token
        self.data = None

        self.init_ui()

    def init_ui(self):

        font = QFont("Microsoft YaHei", 11)
        self.setWindowTitle("Generate ZC415 file")

        screen = QDesktopWidget().screenGeometry()
        self.resize(int(screen.width() / 2.5), int(screen.height() / 2.5))
        frame_geo = self.frameGeometry()
        frame_geo.moveCenter(QDesktopWidget().availableGeometry().center())
        self.move(frame_geo.topLeft())

        # 外层布局（用于居中）
        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.setAlignment(Qt.AlignCenter)

        # 创建 QScrollArea，并设置属性
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)  # 让 scroll area 根据内容自适应
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)  # 横向不要滚动条（可选）
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)  # 纵向滚动条按需出现

        # 内层容器和布局（承载内容）
        inner_widget = QWidget()
        inner_layout = QVBoxLayout(inner_widget)
        inner_layout.setContentsMargins(100, 50, 100, 50)
        inner_layout.setSpacing(20)
        inner_layout.setAlignment(Qt.AlignTop | Qt.AlignHCenter)  # 控制内容居中靠上

        airway_bill = QLabel("Enter AirWayBill:")
        airway_bill.setFont(font)
        self.airwaybill_input = QLineEdit()
        self.airwaybill_input.setFont(font)
        self.airwaybill_input.setMaximumWidth(400)
        self.airwaybill_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

        inner_layout.addWidget(airway_bill)
        inner_layout.addWidget(self.airwaybill_input)

        # download options
        download_option_label = QLabel("Select Download Option:")
        download_option_label.setFont(font)
        self.combo = QComboBox()
        self.combo.addItems(["PDF + XML","PDF only", "XML only"])
        self.combo.setFont(font)
        self.combo.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        inner_layout.addWidget(download_option_label)
        inner_layout.addWidget(self.combo)

        # 底部按钮布局
        inner_layout.addStretch()
        # 底部的确认和取消按钮
        bottom_layout = QHBoxLayout()
        self.ok_btn = QPushButton("OK")
        self.ok_btn.setDefault(True)  # 设置为默认按钮
        self.ok_btn.clicked.connect(self.download_zc415)  # 点击确认按钮
        self.cancel_btn = QPushButton("Close")
        self.cancel_btn.clicked.connect(self.reject)  # 关闭窗口
        bottom_layout.addWidget(self.ok_btn)
        bottom_layout.addWidget(self.cancel_btn)
        inner_layout.addLayout(bottom_layout)

        scroll_area.setWidget(inner_widget)
        # 把内层部件加入外层居中布局
        outer_layout.addWidget(scroll_area)

    def download_zc415(self):

        try:
            # 是否要支持传多个
            air_way_bill = self.airwaybill_input.text().strip()
            option = self.combo.currentText()

            if not air_way_bill:
                QMessageBox.warning(self, "Input Error", "Please enter the airwaybill.")
                return

            response = download_zc415_file(air_way_bill)

            save_dir, _ = QFileDialog.getSaveFileName(self, "Save ZC415 File", "Select Folder to Save")
            if not save_dir:
                return

            with open(save_dir, 'wb') as f:
                f.write(response.content)

            QMessageBox.information(self, "Success", f"File saved to:\n{save_dir}")
            self.accept()  # 关闭对话框

        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save file.\nError message: {str(e)}")

    # 模拟下载函数：实际中你要替换为调用你的第三方 API
    def download_files(self, airwaybill, option):
        # 示例返回值结构：{filename: filecontent_bytes}
        # 实际可以调用 requests、http.client 等库访问你的后端接口
        # 示例内容：
        files = {}

        if option in ["PDF only", "PDF + XML"]:
            files[f"{airwaybill}.pdf"] = b"%PDF-1.4 fake pdf content"

        if option in ["XML only", "PDF + XML"]:
            files[f"{airwaybill}.xml"] = b"<root><awb>{}</awb></root>".format(airwaybill).encode()

        return files