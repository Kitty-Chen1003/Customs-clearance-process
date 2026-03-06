# views/login_page.py
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QScrollArea, QSizePolicy, \
    QApplication, QHBoxLayout
from PyQt5.QtCore import pyqtSignal, Qt

import global_state


class Login(QWidget):

    login_successful = pyqtSignal()

    cancel_requested = pyqtSignal()

    def __init__(self, api):

        super().__init__()
        self.step_config = None
        self.api = api

        self.username_input = None
        self.password_input = None

        self.init_ui()
    
    def set_step_config(self, step_config: dict):
        """设置字段名和提示"""
        self.step_config = step_config
        self.update_hint_label()

    def update_hint_label(self):
        """在页面标题下方增加提示"""
        if not self.step_config or not hasattr(self, "inner_layout"):
            return

        # 删除已有提示标签（如果存在）
        for i in reversed(range(self.inner_layout.count())):
            widget = self.inner_layout.itemAt(i).widget()
            if isinstance(widget, QLabel) and getattr(widget, "is_hint_label", False):
                self.inner_layout.removeWidget(widget)
                widget.deleteLater()

        hint_text = self.step_config.get("hint", "")
        if hint_text:
            hint_label = QLabel(hint_text)
            hint_label.setObjectName("hint")
            hint_label.setAlignment(Qt.AlignCenter)
            hint_label.is_hint_label = True  # 标记为提示
            # 找到标题位置后插入到其下方
            # 假设标题总是 inner_layout 的第一个 QLabel
            for i in range(self.inner_layout.count()):
                widget = self.inner_layout.itemAt(i).widget()
                if isinstance(widget, QLabel) and widget.objectName() == "title":
                    self.inner_layout.insertWidget(i + 1, hint_label)
                    break

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

        # 标题
        title = QLabel("Login")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        inner_layout.addWidget(title)

        inner_layout.setSpacing(20)

        input_prompt_label = QLabel("Please enter the username and password:")
        input_prompt_label.setObjectName("normal")

        inner_layout.addWidget(input_prompt_label)

        inner_layout.setSpacing(20)

        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("username")
        self.username_input.setFocus()
        self.username_input.setMaximumWidth(800)
        self.username_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        inner_layout.addWidget(self.username_input)

        inner_layout.setSpacing(20)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("password")
        self.password_input.setMaximumWidth(800)
        self.password_input.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.password_input.setEchoMode(QLineEdit.Password)

        inner_layout.addWidget(self.password_input)

        inner_layout.setSpacing(20)

        button_layout = QHBoxLayout()

        self.login_button = QPushButton("Login")
        self.login_button.setObjectName("confirmbutton")
        self.login_button.clicked.connect(self.try_login)
        button_layout.addWidget(self.login_button)

        # 取消按钮
        cancel_button = QPushButton("Cancel")
        cancel_button.setObjectName("cancelButton")
        cancel_button.clicked.connect(self.on_cancel)  # 连接到提示框槽函数
        button_layout.addWidget(cancel_button)

        # 将按钮布局添加到主布局中
        inner_layout.addLayout(button_layout)

        # 底部按钮布局
        inner_layout.addStretch()

        scroll_area.setWidget(inner_widget)
        # 把内层部件加入外层居中布局
        outer_layout.addWidget(scroll_area)

    def try_login(self):
        # 禁用按钮，改样式、文字
        self.findChild(QPushButton, "confirmbutton").setEnabled(False)
        self.findChild(QPushButton, "confirmbutton").setText("Processing...")

        QApplication.processEvents()

        try:
            self.get_data()
            if not self.username_input or not self.password_input:
                QMessageBox.warning(self, "Missing Info", "Please enter the username and password.")
                return

            request_response = self.api.login(self.username_input, self.password_input)

            try:
                if request_response.status_code == 200:
                    response_json = request_response.json()
                    print("get response successfully.")
                    if response_json.get('success', False):
                        print("Login successfully.")
                        QMessageBox.information(self, "Success",
                                                "Login successfully!")
                        self.store_data(self.username_input, self.password_input)
                    else:
                        err_msg = response_json.get('error', 'Unknown error')
                        print(f"Login failed: {err_msg}")
                        QMessageBox.critical(self, "Error", f"Login failed! {err_msg}")
                else:
                    print(f"Failed to login. Status code: {request_response.status_code}")
                    print("Raw response text:", request_response.text)
                    QMessageBox.critical(self, "Error",
                                         f"Failed to login. Status code: {request_response.status_code}")
            except AttributeError:
                print("Error: Response object does not have a status_code attribute.")
                QMessageBox.critical(self, "Error", "Invalid response object received.")

        finally:
            # 恢复按钮状态
            self.findChild(QPushButton, "confirmbutton").setEnabled(True)
            self.findChild(QPushButton, "confirmbutton").setText("Login")
            QApplication.processEvents()

    def get_data(self):
        self.username_input = self.username_input.text().strip()
        self.password_input = self.password_input.text().strip()

    def store_data(self, username, password):
        global_state.username = username
        global_state.password = password

    def reset(self):
        # 清除输入框中的文本
        self.username_input.clear()
        self.password_input.clear()

    def on_cancel(self):
        # 创建提示框，确认取消操作
        reply = QMessageBox.question(self, 'Confirm Cancel',
                                     "This action will discard all changes and return to the first page. Are you sure you want to continue?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.cancel_requested.emit()  # 发送信号，通知父窗口取消操作
        else:
            print("User canceled the action.")


