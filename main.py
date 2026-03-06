import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QMessageBox, QStackedWidget, QDesktopWidget, QToolBar, QAction,
    QDialog
)

from company_factory import CompanyFactory
from config_manager import CompanyConfigManager

VIEW_VERSION = "v2"
if VIEW_VERSION == "v2":
    from views.add_awb import AddAwb
    from views.add_arn import AddArn
    from views.add_ens import AddEns
    from views.get_dsk import GetDsk
    from views.add_items_clearances import AddItemsClearances
elif VIEW_VERSION == "v1":
    from views.results.xx.add_awb import AddAwb
    from views.results.xx.add_arn import AddArn
    from views.results.xx.add_ens import AddEns
    from views.results.xx.get_dsk import GetDsk
    from views.results.xx.add_items_clearances import AddItemsClearances
# from views.add_awb import AddAwb
# from views.add_arn import AddArn

from views.add_items_clearances_lsas import AddItemsClearancesLSAS
from views.add_items_releases_lsas import AddItemsReleasesLSAS


from views.add_items_releases import AddItemsReleases
from views.generate_zc415 import GenerateZC415

from views.login import Login
from views.company_selector import CompanySelector
from views.mark_arrived import MarkArrived

STEP_MAP = {
    "Login": Login,
    "AddAwb": AddAwb,
    "AddArn": AddArn,
    "AddEns": AddEns,
    "GetDsk": GetDsk,
    "AddItemsClearances": AddItemsClearances,
    "AddItemsClearancesLSAS": AddItemsClearancesLSAS,
    "AddItemsReleases": AddItemsReleases,
    "AddItemsReleasesLSAS": AddItemsReleasesLSAS,
    "MarkArrived": MarkArrived,
}


class StepWizard(QMainWindow):
    def __init__(self):
        super().__init__()

        # TODO: 和吕总确认名称
        self.config = None
        self.setWindowTitle("Data Forwarding System")
        # self.showMaximized()

        # 设置为屏幕大小的一半并居中
        screen = QDesktopWidget().screenGeometry()
        self.resize(screen.width() // 2, screen.height() // 2)
        frame_geo = self.frameGeometry()
        frame_geo.moveCenter(QDesktopWidget().availableGeometry().center())
        self.move(frame_geo.topLeft())

        self.api = None
        self.steps = []
        self.current_step = 0

        self.create_toolbar()

        self.stack = QStackedWidget()

        self.company_selector = CompanySelector()
        self.company_selector.company_selected.connect(self.on_company_selected)
        self.stack.addWidget(self.company_selector)
        self.steps.append(self.company_selector)

        # self.step1 = AddAwb(self.api)
        # self.step2 = AddArn(self.api)
        # self.step3 = AddEns(self.api)
        # self.step4 = GetDsk(self.api)
        # self.step5 = AddItemsClearance(self.api)
        # self.step6 = AddItemsReleases(self.api)
        # # self.step1.finished_successful.connect(lambda: self.next_button.setEnabled(True))
        # self.step1.cancel_requested.connect(self.reset_all)
        # # self.step2.finished_successful.connect(lambda: self.next_button.setEnabled(True))
        # self.step2.cancel_requested.connect(self.reset_all)
        # # self.step3.finished_successful.connect(lambda: self.next_button.setEnabled(True))
        # self.step3.cancel_requested.connect(self.reset_all)
        # # self.step4.finished_successful.connect(lambda: self.next_button.setEnabled(True))
        # self.step4.cancel_requested.connect(self.reset_all)
        # # self.step5.finished_successful.connect(lambda: self.next_button.setEnabled(True))
        # self.step5.cancel_requested.connect(self.reset_all)
        # # self.step6.finished_successful.connect(lambda: self.next_button.setEnabled(True))
        # self.step6.cancel_requested.connect(self.reset_all)
        # self.stack.addWidget(self.step1)
        # self.stack.addWidget(self.step2)
        # self.stack.addWidget(self.step3)
        # self.stack.addWidget(self.step4)
        # self.stack.addWidget(self.step5)
        # self.stack.addWidget(self.step6)

        # 底部导航按钮
        self.prev_button = QPushButton("Previous")
        self.next_button = QPushButton("Next")

        # self.prev_button.setObjectName("processButton")
        # self.next_button.setObjectName("processButton")

        # self.next_button.setEnabled(False)

        self.prev_button.clicked.connect(self.go_prev)
        self.next_button.clicked.connect(self.go_next)

        self.nav_layout = QHBoxLayout()
        self.nav_layout.addWidget(self.prev_button)
        self.nav_layout.addStretch()
        self.nav_layout.addWidget(self.next_button)

        # 总布局
        self.main_layout = QVBoxLayout()
        self.main_layout.addWidget(self.stack)
        self.main_layout.addLayout(self.nav_layout)

        container = QWidget()
        container.setLayout(self.main_layout)
        self.setCentralWidget(container)

        self.update_buttons()

    # ---------------- 公司选择后动态加载步骤 ----------------
    def on_company_selected(self, company_name: str):
        self.api = CompanyFactory.get_api(company_name)
        self.config = CompanyConfigManager()  # 新增

        # 清理旧步骤（保留 company_selector）
        while self.stack.count() > 1:
            widget = self.stack.widget(1)
            self.stack.removeWidget(widget)
            widget.deleteLater()

        self.steps = [self.company_selector]

        # Excel 驱动步骤
        for step_name in self.config.get_steps(company_name):
            step_class = STEP_MAP.get(step_name)
            if not step_class:
                continue

            page = step_class(self.api)
            # 传入 step 配置
            if hasattr(page, "set_step_config"):
                step_config = self.config.get_step_config(company_name, step_name)
                page.set_step_config(step_config)

            page.cancel_requested.connect(self.reset_all)
            self.stack.addWidget(page)
            self.steps.append(page)

        # 跳转到第一步
        self.current_step = 1
        self.update_buttons()

    def create_toolbar(self):
        toolbar = QToolBar("ToolBar", self)
        toolbar.setStyleSheet("""
            QToolButton {
                background-color: #0078d7;
                color: white;
                border: 1px solid #005a9e;
                border-radius: 4px;
                padding: 6px 12px;
            }
            QToolButton:hover {
                background-color: #005a9e;
            }
            QToolButton:pressed {
                background-color: #003f7f;
            }
        """)
        self.addToolBar(toolbar)

        # 创建现有工具栏按钮
        self.toolbar_download_ZC415 = QAction('Generate ZC415', self)

        # 将现有按钮添加到工具栏中
        toolbar.addSeparator()
        toolbar.addAction(self.toolbar_download_ZC415)
        toolbar.addSeparator()  # 添加分隔符

        # 连接新按钮的触发信号

        self.toolbar_download_ZC415.triggered.connect(self.generate_zc415)

    def update_buttons(self):
        self.stack.setCurrentIndex(self.current_step)
        self.prev_button.setEnabled(self.current_step > 0)
        if self.current_step == self.stack.count() - 1:
            self.next_button.setText("Finish")
        else:
            self.next_button.setText("Next")

        # 若当前页面提交过，才启用 next
        # current_widget = self.stack.currentWidget()
        # if hasattr(current_widget, 'finished') and current_widget.finished:
        #     self.next_button.setEnabled(True)
        # else:
        #     self.next_button.setEnabled(False)

    def go_prev(self):
        if self.current_step > 0:
            self.current_step -= 1
            self.update_buttons()

    def go_next(self):
        # 如果当前不是最后一步，前往下一步
        if self.current_step < self.stack.count() - 1:
            self.current_step += 1
            self.update_buttons()
        else:
            # 如果是最后一步，重置为第一步
            QMessageBox.information(self, "Info", "Finished, all data submitted successful!")
            # self.current_step = 0
            # self.update_buttons()
            self.reset_all()
        # self.next_button.setEnabled(False)

    def reset_all(self):
        self.current_step = 0
        self.stack.setCurrentIndex(0)
        # self.next_button.setEnabled(False)
        self.update_buttons()

        # 调用每个页面的 reset() 方法清空数据
        # for i in range(1, self.stack.count()):
        #     widget = self.stack.widget(i)
        for step in self.steps[1:]:
            if hasattr(step, "reset"):
                step.reset()

    def generate_zc415(self):
        # dialog = GenerateZC415(self.username, self.token)
        dialog = GenerateZC415()
        if dialog.exec_() == QDialog.Accepted:
            print("对话框成功关闭")
        else:
            print("对话框被取消")
        pass


if __name__ == "__main__":
    app = QApplication(sys.argv)
    with open("style.qss", "r") as f:
        app.setStyleSheet(f.read())
    window = StepWizard()
    window.show()
    sys.exit(app.exec_())
