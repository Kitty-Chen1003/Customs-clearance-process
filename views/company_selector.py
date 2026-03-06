from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QPushButton, QScrollArea
from PyQt5.QtCore import pyqtSignal, Qt


class CompanySelector(QWidget):

    company_selected = pyqtSignal(str)

    def __init__(self):

        super().__init__()

        self.init_ui()

    def init_ui(self):

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

        title = QLabel("Choose company")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        inner_layout.addWidget(title)

        # 你可以做成按钮组，或者下拉框
        ecargo_btn = QPushButton("Ecargo")
        ecargo_btn.setObjectName("companyButton")
        ecargo_btn.clicked.connect(lambda: self.company_selected.emit("Ecargo"))

        lsas_btn = QPushButton("LSAS")
        lsas_btn.setObjectName("companyButton")
        lsas_btn.clicked.connect(lambda: self.company_selected.emit("LSAS"))

        inner_layout.addWidget(ecargo_btn)
        inner_layout.addWidget(lsas_btn)

        inner_layout.addStretch()
        scroll_area.setWidget(inner_widget)
        # 把内层部件加入外层居中布局
        outer_layout.addWidget(scroll_area)
