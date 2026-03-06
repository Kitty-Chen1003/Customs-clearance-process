from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QLineEdit, QSizePolicy, QPushButton, QHBoxLayout, QFileDialog, \
    QMessageBox, QScrollArea, QApplication

import global_state
from http_client import add_ens_request


class MarkArrived(QWidget):

    # finished_successful = pyqtSignal()
    cancel_requested = pyqtSignal()

    def __init__(self, api):

        super().__init__()
        self.step_config = None
        self.api = api

        # self.finished = False

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
        title = QLabel("Mark Arrived")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)
        inner_layout.addWidget(title)

        inner_layout.addSpacing(50)

        # # 文件上传按钮
        # upload_label = QLabel("upload process file:")
        # upload_label.setObjectName("normal")
        # inner_layout.addWidget(upload_label)
        #
        # upload_button = QPushButton("Select file")
        # upload_button.setFont(QFont("Microsoft YaHei", 12))
        # upload_button.setMaximumWidth(800)
        # upload_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        # upload_button.setStyleSheet("padding: 10px;")
        # upload_button.clicked.connect(self.select_file)
        #
        # # 创建 QLabel 并放入 QScrollArea 中
        # self.label_status = QLabel('No file selected')
        # self.label_status.setWordWrap(True)
        # self.label_status.setObjectName("normal")
        # self.scroll_area_label = QScrollArea()
        # self.scroll_area_label.setMaximumWidth(800)
        # self.scroll_area_label.setWidgetResizable(True)
        # self.scroll_area_label.setMinimumHeight(100)
        # self.scroll_area_label.setWidget(self.label_status)
        #
        # inner_layout.addWidget(upload_button)
        # inner_layout.addSpacing(20)
        # inner_layout.addWidget(self.scroll_area_label)
        # inner_layout.addSpacing(20)

        # 按钮水平布局（发送按钮和取消按钮并排）
        button_layout = QHBoxLayout()

        # 发送按钮
        submit_button = QPushButton("Submit")
        submit_button.setObjectName("confirmbutton")
        submit_button.clicked.connect(self.submit)
        submit_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        button_layout.addWidget(submit_button)

        # 取消按钮
        cancel_button = QPushButton("Cancel")
        cancel_button.setObjectName("cancelButton")
        cancel_button.clicked.connect(self.on_cancel)  # 连接到提示框槽函数
        cancel_button.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        button_layout.addWidget(cancel_button)

        # 将按钮布局添加到主布局中
        inner_layout.addLayout(button_layout)
        inner_layout.addSpacing(20)

        # clear_button = QPushButton("Clear")
        # clear_button.setFont(QFont("Microsoft YaHei", 12))
        # clear_button.setStyleSheet("padding: 10px;")
        # clear_button.clicked.connect(self.reset)
        # inner_layout.addWidget(clear_button)

        inner_layout.addStretch()

        scroll_area.setWidget(inner_widget)
        # 把内层部件加入外层居中布局
        outer_layout.addWidget(scroll_area)

    # def reset(self):
    #     # 重置文件路径和上传按钮文本
    #     self.file_path = ""
    #     self.label_status.setText("No file selected")

    def submit(self):
        # 禁用按钮，改样式、文字
        self.findChild(QPushButton, "confirmbutton").setEnabled(False)
        self.findChild(QPushButton, "confirmbutton").setText("Processing...")

        QApplication.processEvents()

        try:

            request_response = self.api.mark_arrived()
            # self.finished = True
            # self.finished_successful.emit()
            try:
                if request_response.status_code == 200:
                    response_json = request_response.json()
                    print("get response successfully.")
                    if response_json.get('success', False):
                        print("Mark arrived successfully.")
                        QMessageBox.information(self, "Success", "Mark arrived successfully!")
                        # self.finished = True
                        # self.finished_successful.emit()  # 发送信号，通知父窗口更新数据
                    else:
                        err_msg = response_json.get('error', 'Unknown error')
                        print(f"Mark arrived failed: {err_msg}")
                        QMessageBox.critical(self, "Error", f"Mark arrived failed! {err_msg}")

                else:
                    print(f"Failed to mark arrived. Status code: {request_response.status_code}")
                    print("Raw response text:", request_response.text)
                    QMessageBox.critical(self, "Error", f"Failed to mark arrived. Status code: {request_response.status_code}")
            except AttributeError:
                print("Error: Response object does not have a status_code attribute.")
                QMessageBox.critical(self, "Error", "Invalid response object received.")

        finally:
            # 恢复按钮状态
            self.findChild(QPushButton, "confirmbutton").setEnabled(True)
            self.findChild(QPushButton, "confirmbutton").setText("Submit")
            QApplication.processEvents()

    def on_cancel(self):
        # 创建提示框，确认取消操作
        reply = QMessageBox.question(self, 'Confirm Cancel',
                                     "This action will discard all changes and return to the first page. Are you sure you want to continue?",
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            self.cancel_requested.emit()  # 发送信号，通知父窗口取消操作
        else:
            print("User canceled the action.")

    # def select_file(self):
    #
    #     file_dialog = QFileDialog(self)
    #     file_dialog.setWindowTitle('Select Process File')
    #     file_dialog.setFileMode(QFileDialog.ExistingFile)  # 只允许选择一个文件
    #     file_dialog.setNameFilter("Excel Files (*.xls *.xlsx)")
    #
    #     if file_dialog.exec_():
    #         selected_file = file_dialog.selectedFiles()[0]  # 只取第一个
    #         self.file_path = selected_file  # 保存文件路径
    #         self.label_status.setText(self.file_path)  # 显示文件路径