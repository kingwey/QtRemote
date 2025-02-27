import socket
import sys
import winreg

import ldap3
import win32security
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox, QCheckBox
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt


class LoginPage(QWidget):

    def __init__(self):
        super().__init__()

        self.init_ui()
        self.detail_page = None

    def init_ui(self):
        # 设置窗口标题和大小
        self.setWindowTitle('登录')
        # 设置固定大小，禁止放大和缩小
        self.setFixedSize(400, 300)

        # 设置背景颜色和字体
        self.setStyleSheet("background-color: #2c3e50; color: white;")
        font = QFont('Arial', 12)

        # 创建标题
        title = QLabel('欢迎登录', self)
        title.setFont(QFont('Arial', 24))
        title.setStyleSheet("color: #ecf0f1;")
        title.setAlignment(Qt.AlignCenter)

        # 创建用户名标签和输入框
        username_label = QLabel('用户名', self)
        username_label.setFont(font)
        self.username_edit = QLineEdit(self)
        self.username_edit.setFont(font)
        self.username_edit.setStyleSheet("background-color: #34495e; color: white; padding: 5px; border-radius: 5px;")

        # 创建密码标签和输入框
        password_label = QLabel('密码', self)
        password_label.setFont(font)
        self.password_edit = QLineEdit(self)
        self.password_edit.setFont(font)
        self.password_edit.setEchoMode(QLineEdit.Password)
        self.password_edit.setStyleSheet("background-color: #34495e; color: white; padding: 5px; border-radius: 5px;")

        # 创建显示密码的复选框
        self.show_password_checkbox = QCheckBox('显示密码', self)
        self.show_password_checkbox.setFont(font)
        self.show_password_checkbox.setStyleSheet("color: #ecf0f1;")
        self.show_password_checkbox.stateChanged.connect(self.toggle_password_visibility)

        # 创建登录按钮
        login_button = QPushButton('登录', self)
        login_button.setFont(font)
        login_button.setStyleSheet("background-color: #e74c3c; color: white; padding: 10px; border-radius: 5px;")
        login_button.clicked.connect(self.login)

        # 创建垂直布局并添加控件
        vbox = QVBoxLayout()
        vbox.addWidget(title)
        vbox.addStretch(1)
        vbox.addWidget(username_label)
        vbox.addWidget(self.username_edit)
        vbox.addWidget(password_label)
        vbox.addWidget(self.password_edit)
        vbox.addWidget(self.show_password_checkbox)
        vbox.addStretch(1)
        vbox.addWidget(login_button)

        # 设置窗口主布局
        self.setLayout(vbox)

    def toggle_password_visibility(self):
        # 切换密码显示模式
        if self.show_password_checkbox.isChecked():
            self.password_edit.setEchoMode(QLineEdit.Normal)
        else:
            self.password_edit.setEchoMode(QLineEdit.Password)

    def login(self):
        # 获取输入的用户名和密码
        username = self.username_edit.text()
        password = self.password_edit.text()

        if self.authenticate_domain_user(username, password):
            self.show_detail_page("域账号登录")
        elif self.authenticate_local_user(username, password):
            self.show_detail_page("本地用户登录")
        elif username == 'admin' and password == 'password123':
            self.show_detail_page("超级管理员登录")
        else:
            QMessageBox.warning(self, '错误', '用户名或密码错误。')

    def show_detail_page(self, login_type="用户登录"):
        if self.detail_page is None:
            self.detail_page = DetailPage(login_type)
        self.detail_page.show()
        self.close()

    @classmethod
    # 验证本地用户的账号密码
    def authenticate_local_user(cls, username, password):
        """验证本地用户的账号密码"""
        try:
            # 获取本地计算机名称
            computer_name = socket.gethostname()

            # 验证密码
            token = win32security.LogonUser(
                username,
                computer_name,
                password,
                win32security.LOGON32_LOGON_INTERACTIVE,
                win32security.LOGON32_PROVIDER_DEFAULT
            )

            # 验证成功
            return True

        except Exception as e:
            # 验证失败
            print(f"Authentication failed: {e}")
            return False

    @classmethod
    # 验证域用户的账号密码
    def authenticate_domain_user(cls, username, password):
        """验证域用户的账号密码"""
        try:
            # 获取注册表中的域名
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Services\Tcpip\Parameters")
            domain, reg_type = winreg.QueryValueEx(key, 'Domain')
            key.Close()

            # 验证用户名密码
            server = ldap3.Server(f'ldap://{domain}', get_info=ldap3.ALL)
            with ldap3.Connection(server, user=f'{username}@{domain}', password=password, auto_bind=True) as conn:
                if conn.bound:
                    conn.unbind()
                    return True
                else:
                    return False

        except Exception as e:
            print(f"LDAP authentication exception: {e}")
            return False


class DetailPage(QWidget):
    def __init__(self, login_type):
        self.login_type = login_type
        super().__init__()

        self.init_ui()

    def init_ui(self):
        # 设置窗口标题和大小
        self.setWindowTitle('详情页面')
        self.setFixedSize(400, 300)

        # 设置背景颜色和字体
        self.setStyleSheet("background-color: #2c3e50; color: white;")
        font = QFont('Arial', 12)

        # 创建标题
        title = QLabel(f'详情内容：{self.login_type}', self)
        title.setFont(QFont('Arial', 24))
        title.setStyleSheet("color: #ecf0f1;")
        title.setAlignment(Qt.AlignCenter)

        # 创建垂直布局并添加控件
        vbox = QVBoxLayout()
        vbox.addStretch(1)
        vbox.addWidget(title)
        vbox.addStretch(1)

        # 设置窗口主布局
        self.setLayout(vbox)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    login_page = LoginPage()
    login_page.show()
    sys.exit(app.exec_())
