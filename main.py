import csv
import ctypes
import getpass
import sys
import time

from PyQt5 import QtCore, uic, QAxContainer, QtWidgets
from PyQt5.QtCore import QTimer, Qt
from PyQt5.QtGui import QIcon
from PyQt5.QtNetwork import QLocalSocket, QLocalServer
from PyQt5.QtWidgets import QMainWindow, QApplication, QHeaderView, QTableWidgetItem, QMenu, QMessageBox, \
    QAbstractItemView, QFileDialog, QSystemTrayIcon, QAction, QWidget

from core.common import print, HostRemote

DISC_REASON = {
    "1": "本地断开连接, 重新链接",
    "2308": "套接字已关闭",
    "3": "按服务器远程断开连接",
    "3080": "解压缩错误",
    "264": "连接超时",
    "3078": "解密错误",
    "260": "DNS 名称查找失败",
    "1288": "DNS 查找失败",
    "2822": "加密错误",
    "1540": "Windows 套接字 gethostbyname 调用失败",
    "520": "找不到主机错误",
    "1032": "内部错误",
    "2310": "内部安全错误",
    "2566": "内部安全错误",
    "1286": "指定的加密方法无效",
    "2052": "指定的 IP 地址不正确",
    "1542": "服务器安全数据无效",
    "1030": "安全数据无效",
    "776": "指定的 IP 地址无效",
    "2056": "许可证协商失败",
    "2312": "许可超时",
    "0": "没有可用信息",
    "262": "内存不足",
    "518": "内存不足",
    "774": "内存不足",
    "2": "用户远程断开连接 这不是错误代码",
    "1798": "未能解包服务器证书",
    "516": "Windows 套接字 连接 失败",
    "1028": "Windows 套接字 recv 调用失败",
    "1796": "发生超时",
    "1544": "内部计时器错误",
    "772": "Windows 套接字 发送 调用失败",
    "2823": "帐户已禁用",
    "3591": "帐户已过期",
    "3335": "帐户已锁定",
    "3079": "帐户受限",
    "6919": "收到的证书已过期",
    "5639": "该策略不支持将凭据委派到目标服务器",
    "8455": "服务器身份验证策略不允许使用保存的凭据进行连接请求 用户必须输入新凭据",
    "2055": "登录失败",
    "6151": "无法联系任何机构进行身份验证 身份验证方的域名可能不正确，域无法访问，或者可能存在信任关系失败",
    "2567": "指定的用户没有帐户",
    "3847": "密码已过期",
    "4615": "首次登录之前，必须更改用户密码",
    "5895": "除非已实现相互身份验证，否则不允许将凭据委派给目标服务器",
    "8711": "智能卡被阻止",
    "7175": "向智能卡显示错误的 PIN"
}
I_ERROR = {
    "-5": "会话争用",
    "-2": "将继续执行登录过程",
    "-3": "正在默默结束",
    "-6": "无权限",
    "-7": "断开连接被拒绝",
    "-4": "重新连接",
    "-1": "用户被拒绝访问",
    "0": "登录失败，因为登录凭据无效",
    "2": "出现另一个登录或登录后错误，远程桌面客户端向用户显示登录屏幕",
    "1": "密码已过期，用户必须更新其密码才能继续登录",
    "3": "远程桌面客户端显示一个对话框，其中包含用户的重要信息",
    "-1073741714": "用户名和身份验证信息有效，但由于用户帐户的限制（如一天中的时间限制），身份验证被阻止",
    "-1073741715": "尝试的登录无效，这是因为用户名不正确或身份验证信息不正确",
    "-1073741276": "密码已过期，用户必须更新其密码才能继续登录"
}
ERROR_CODE = {
    "0": "未知错误",
    "1": "内部错误代码 1",
    "2": "内存不足",
    "3": "窗口创建错误",
    "4": "内部错误代码 2",
    "5": "内部错误代码 3； 这不是有效状态",
    "6": "内部错误代码 4",
    "7": "在客户端连接期间发生了不可恢复的错误",
    "100": "Winsock 初始化错误"
}


class Main(QMainWindow, HostRemote):

    def __init__(self):
        # 超类
        super(QMainWindow, self).__init__()
        super(HostRemote, self).__init__()

        # 加载ui文件
        self.ui = uic.loadUi(r'ui\main.ui', self)

        # 打印日志至窗体控件
        try:
            sys.stdout.write = self.log_plain
        except Exception as ex:
            print(ex)

        # 判断管理员权限
        if ctypes.windll.shell32.IsUserAnAdmin():
            print("管理员权限启动应用，可以连接远程桌面了！")
        else:
            print("用户权限启动应用，可以连接远程桌面了！")

        # 引入全部样式
        with open(r"css\main.css", encoding="utf-8") as f:
            self.setStyleSheet(f.read())

        # 系统托盘
        self.system_tray_icon = self.system_tray()
        # 显示托盘
        self.system_tray_icon.show()

        # 设置窗口标题
        self.setWindowTitle(self.window_title)

        # 不显示行序号
        self.table_desk.verticalHeader().setVisible(False)

        # 列宽自适应
        self.table_desk.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 行高自适应
        # self.table_desk.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 隔行交替背景色
        self.table_desk.setAlternatingRowColors(True)

        # 设置选中时为整行选中
        self.table_desk.setSelectionBehavior(QAbstractItemView.SelectRows)

        # 初始化表格内容
        self.init_table()

        # 绑定表格右键事件
        self.table_desk.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.table_desk.customContextMenuRequested[QtCore.QPoint].connect(self.table_right_click_event)

        # 双击最大化
        self.table_desk.itemDoubleClicked.connect(self.item_double_clicked_event)

        # 表格排序
        self.table_desk.horizontalHeader().sectionClicked.connect(self.sort_table_event)

        # 表格数据更新时的事件
        self.table_desk.itemChanged.connect(self.item_changed_event)

        # 绑定添加按钮的事件
        self.button_append.clicked.connect(self.append_desk)

        self.edit_webhook.setText(self.get_preferences_keys(["notice", "webhook"]))
        self.edit_secret.setText(self.get_preferences_keys(["notice", "secret"]))
        self.edit_per.setText(self.get_preferences_keys(["notice", "per"]))

        # webhook secret per
        self.edit_webhook.editingFinished.connect(self.webhook_editing_finished)
        self.edit_secret.editingFinished.connect(self.secret_editing_finished)
        self.edit_per.editingFinished.connect(self.per_editing_finished)

        # 定时器
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.check_remote)
        self.timer.start(1000 * 20)

    def webhook_editing_finished(self):
        webhook = self.edit_webhook.text()
        self.set_preferences_keys(["notice", "webhook"], value=webhook)

    def secret_editing_finished(self):
        secret = self.edit_secret.text()
        self.set_preferences_keys(["notice", "secret"], value=secret)

    def per_editing_finished(self):
        per = self.edit_per.text()
        self.set_preferences_keys(["notice", "per"], value=per)

    def on_tray_icon_activated(self, reason):
        """系统托盘点击显示窗口"""
        if reason == QSystemTrayIcon.Trigger:
            self.restore_window()

    def restore_window(self):
        """恢复窗口"""
        if self.isMinimized():
            self.showMaximized()
        else:
            self.show()
        self.activateWindow()

    def system_tray(self):
        """系统托盘"""
        # 系统托盘图标
        system_tray_icon = QSystemTrayIcon(QIcon(r'img\logo.ico'), self)
        # 设置悬浮提示
        system_tray_icon.setToolTip(self.window_title)
        # 鼠标点击
        system_tray_icon.activated.connect(self.on_tray_icon_activated)

        # 菜单
        menu = QMenu()

        # 打开主页
        go_home = QAction(QIcon(r"img\home.svg"), '打开主界面', self)
        go_home.triggered.connect(self.restore_window)
        menu.addAction(go_home)

        menu.addSeparator()

        # 退出
        logout = QAction(QIcon(r"img\logout.png"), '退出', self)
        logout.triggered.connect(QApplication.instance().quit)
        menu.addAction(logout)

        system_tray_icon.setContextMenu(menu)

        return system_tray_icon

    @classmethod
    def rdp_control(cls):
        """QAxWidget RDP 控件"""
        ax_rdp = QAxContainer.QAxWidget()
        # 设置com activeX控件
        ax_rdp.setControl("{8B918B82-7985-4C24-89DF-C33AD2BBFBCD}")
        # ##################################属性设置################################### #
        ax_rdp.setProperty("enabled", False)

        ax_rdp.setObjectName("ax_rdp")

        # 一般属性
        ax_rdp.setProperty("DesktopWidth", 1920)
        ax_rdp.setProperty("DesktopHeight", 1080)
        # 不显示任何警告信息
        ax_rdp.setProperty("DisplayAlerts", False)
        # 显示滚动条
        ax_rdp.setProperty("DisplayScrollBars", True)

        # 高级属性
        advanced_settings7 = ax_rdp.querySubObject("AdvancedSettings7")

        # 端口
        advanced_settings7.setProperty("RDPPort", 3389)
        # 显示连接栏
        advanced_settings7.setProperty("DisplayConnectionBar", True)
        # 连接栏显示最小化按钮
        advanced_settings7.setProperty("ConnectionBarShowMinimizeButton", False)
        # 连接栏显示还原按钮
        advanced_settings7.setProperty("ConnectionBarShowRestoreButton", True)
        # 连接栏显示固定按钮
        advanced_settings7.setProperty("ConnectionBarShowPinButton", False)
        # 智能调整大小
        advanced_settings7.setProperty("SmartSizing", True)
        # 启用凭据安全服务提供程序 (CredSSP)
        advanced_settings7.setProperty("EnableCredSspSupport", True)
        # 重定向智能卡
        advanced_settings7.setProperty("RedirectSmartCards", True)
        # 连接栏固定按钮状态
        advanced_settings7.setProperty("PinConnectionBar", True)
        # 启用压缩,减小带宽
        advanced_settings7.setProperty("Compress", 1)
        # 剪切板
        advanced_settings7.setProperty("RedirectClipboard", True)
        # 禁止：自动重新连接
        advanced_settings7.setProperty("CanAutoReconnect", False)
        advanced_settings7.setProperty("EnableAutoReconnect", False)
        # 是否应显示重定向警告对话框
        advanced_settings7.setProperty("ShowRedirectionWarningDialog", False)
        # 不对服务器进行身份验证
        advanced_settings7.setProperty("AuthenticationLevel", 0)
        return ax_rdp

    def init_table(self):
        """初始化表格内容"""
        # 清空所有行
        while self.table_desk.rowCount() > 0:
            self.table_desk.removeRow(0)

        # 初始化
        for desk_info in self.host_info:
            self.add_row(desk_info)

    def add_row(self, desk_info):
        """新建表格控件行对象，并添加至表格"""
        try:
            if len(desk_info[0]) == 0 or len(desk_info[1]) == 0 or len(desk_info[2]) == 0 or len(desk_info[3]) == 0:
                print("添加桌面：有不完整的内容", log_type="警告")
                QMessageBox.information(self, "新增失败", "有不完整的内容", QMessageBox.Yes)
                return

            row = self.table_desk.rowCount()

            # 新增一行
            self.table_desk.insertRow(row)
            self.table_desk.setRowHeight(row, 108)
            # 添加桌面信息在至表格
            for index, info in enumerate(desk_info):
                item = QTableWidgetItem(str(info))
                item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)

                if index == 2:
                    # 隐藏密码
                    item.setText("*" * len(info))
                    item.setData(QtCore.Qt.UserRole, info)

                self.table_desk.setItem(row, index, item)

            # 添加状态项目至table
            status_item = QTableWidgetItem("未连接")
            status_item.setData(QtCore.Qt.UserRole, "")  # 存储错误信息
            status_item.setData(QtCore.Qt.UserRole + 1, 0)  # 存储错误次数
            status_item.setTextAlignment(QtCore.Qt.AlignHCenter | QtCore.Qt.AlignVCenter)
            self.table_desk.setItem(row, 4, status_item)

            # 添加rdp控件至table
            ax_rdp = self.rdp_control()
            self.table_desk.setCellWidget(row, 5, ax_rdp)

            # 刷新表格
            self.table_desk.update()
        except Exception as ex:
            print("add_row", str(ex), log_type="错误")

    def log_plain(self, text):
        """打印日志至窗体控件"""
        try:
            self.plain_text_log.appendPlainText(text.strip())
            time.sleep(0.05)
        except Exception as ex:
            _ = ex

    # ********控件:事件处理*********
    def closeEvent(self, event):
        """关闭窗口事件"""
        # event.accept()

        # 禁止窗口关闭，隐藏窗口
        event.ignore()
        self.hide()

    def append_desk(self):
        """添加桌面信息"""
        # 桌面信息
        host = self.edit_address.text()
        username = self.edit_account.text()
        password = self.edit_password.text()
        domain = self.edit_domain.text()

        desk_info = [host, username, password, domain]
        if self.select(desk_info):
            print("添加桌面：远程桌面已存在", log_type="警告")
            QMessageBox.information(self, "新增失败", "远程桌面已存在", QMessageBox.Yes)
            return

        self.add_row(desk_info)
        if len(desk_info[0]) == 0 or len(desk_info[1]) == 0 or len(desk_info[2]) == 0 or len(desk_info[3]) == 0:
            return
        # 缓存至本地文件
        self.add(desk_info)
        print(f"添加桌面 {host} {username}")

    def sort_table_event(self, column):
        """表格排序"""
        print("点击表头【排序】")
        self.table_desk.sortItems(column, Qt.SortOrder.AscendingOrder)

    def item_changed_event(self, item):
        """表格项目信息变化"""
        try:
            row = item.row()
            column = item.column()
            text = item.text()

            # 限制前3列数据可以编辑
            if column > 3:
                return

            # 密码
            if column == 2:
                password = item.data(QtCore.Qt.UserRole)

                if all(i == "*" for i in text):
                    if len(text) == len(password):
                        return
                    text = password
                elif text is None or text == "":
                    text = password
                else:
                    item.setData(QtCore.Qt.UserRole, text)

            # 行数不能超过最大行号
            if row >= len(self.host_info) - 1:
                return

            # 同步修改数据缓存到本地
            if text != self.host_info[row][column]:
                print(item.row(), "行", item.column(), "列", "项目数据发生变化")
                self.host_info[row][column] = text
                self.modify(row, self.host_info[row])

            # 隐藏密码
            if column == 2:
                item.setText("*" * len(text))
        except Exception as ex:
            print("表格数据变化", str(ex), log_type="错误")

    def table_right_click_event(self, pos):
        """右键表格事件"""
        try:
            # 弹出菜单
            pop_menu = QMenu()
            conn_desk = pop_menu.addAction(QIcon("img/connect.svg"), '默认连接')
            pop_menu.addSeparator()
            max_desk = pop_menu.addAction(QIcon("img/full_screen.svg"), '全屏显示')
            pop_menu.addSeparator()
            disconn_desk = pop_menu.addAction(QIcon("img/disconnect.svg"), '断开连接')
            pop_menu.addSeparator()
            del_desk = pop_menu.addAction(QIcon("img/delete.svg"), '删除桌面')
            pop_menu.addSeparator()
            load_file = pop_menu.addAction(QIcon("img/import.svg"), '导入文件')

            # 获取右键菜单中当前被点击的是哪一项
            action = pop_menu.exec_(self.table_desk.mapToGlobal(pos))

            # 获取当前选中的行
            rows = set()
            for item in self.table_desk.selectedItems():
                rows.add(item.row())

            # 将选定行的行号降序排序
            rows = list(rows)
            rows.sort(reverse=True)

            # 导入文件
            if action == load_file:

                file_info, ext = QFileDialog.getOpenFileName(self, "选择文件", self.user_file, "CSV(*.csv)")
                if not file_info:
                    return
                print("操作【导入文件】", file_info)
                with open(file_info, mode='r', newline='', encoding="ANSI", errors='ignore') as f:
                    # csv读取器
                    reader = csv.reader(f, delimiter=",")
                    datas = list(reader)
                    if datas[0] == ["地址", "用户名", "密码", "域名"]:
                        # 重置self.host_info
                        self.host_info = []
                        for i in datas[1:]:
                            if len(i) == 4 and i not in self.host_info:
                                self.host_info.append(i)
                        self.init_table()
                        self.dump()
                    else:
                        QMessageBox.information(self, "导入失败", "表头不正确\n[地址, 用户名, 密码, 域名]",
                                                QMessageBox.Yes)
                return

            # 说明没有选中任何行
            if len(rows) == 0:
                print("没有选中任何行", log_type="警告")
                return

            # 其他操作
            for row in rows:
                # 获取当前选中行的会话信息
                host = self.table_desk.item(row, 0).text()
                username = self.table_desk.item(row, 1).text()
                password = self.table_desk.item(row, 2).data(QtCore.Qt.UserRole)
                domain = self.table_desk.item(row, 3).text()
                status_item = self.table_desk.item(row, 4)
                ax_rdp = self.table_desk.cellWidget(row, 5)
                host_info = [host, username, password, domain]
                if action:
                    # 操作
                    print(host, username, "操作【", action.text(), "】")

                if action == conn_desk:
                    # 选中连接桌面
                    self.connect_event(host, username, password, domain, ax_rdp, status_item)

                elif action == max_desk:
                    # 选中全屏显示
                    if ax_rdp.property("Connected") == 1:
                        ax_rdp.setProperty("FullScreen", False)
                        ax_rdp.setProperty("FullScreen", True)
                    else:
                        print(host, username, "远程桌面未连接", log_type="警告")

                elif action == disconn_desk:
                    # 选中断开连接
                    if ax_rdp.property("Connected") == 1:
                        status_item.setText("断开连接")
                        ax_rdp.dynamicCall("Disconnect()")
                    else:
                        print(host, username, "远程桌面未连接", log_type="警告")

                elif action == del_desk:
                    reply = QMessageBox.question(self, '提示', f"确定要{host} {username}删除桌面吗？",
                                                 QMessageBox.Yes | QMessageBox.No)
                    if reply == QMessageBox.Yes:
                        # 选中删除桌面
                        self.table_desk.removeRow(row)
                        self.delete(host_info)
                        print(f"{host} {username}桌面已删除！")
                    else:
                        print(f"{host} {username}已取消删除桌面。")

                # 刷新表格
                self.table_desk.update()
        except Exception as ex:
            print(f"table_right_click_event {str(ex)}", log_type="错误")

    def item_double_clicked_event(self, item):
        """双击最大化"""
        try:
            row = item.row()
            column = item.column()
            if column < 4:
                return
            host = self.table_desk.item(row, 0).text()
            username = self.table_desk.item(row, 1).text()
            ax_rdp = self.table_desk.cellWidget(row, 5)
            print(host, username, "双击操作【 全屏显示 】")
            # 选中全屏显示
            if ax_rdp.property("Connected") == 1:
                print("测试")

                qw = QWidget()
                qw.horizontalLayout = QtWidgets.QHBoxLayout(qw)
                qw.horizontalLayout.setObjectName("horizontalLayout")
                # QAxWidget控件添加至布局
                qw.horizontalLayout.addWidget(ax_rdp)

                qw.horizontalLayout.setContentsMargins(0, 0, 0, 0)
                qw.showFullScreen()

                # ax_rdp.setProperty("FullScreen", False)
                # ax_rdp.setProperty("FullScreen", True)
            else:
                print(host, username, "远程桌面未连接", log_type="警告")
        except Exception as ex:
            print("表格数据变化", str(ex), log_type="错误")

    # ********默认连接******** #
    @classmethod
    def connect_event(cls, host, username, password, domain, ax_rdp, status_item):

        # 打印日志
        def log(text="", status="登录错误", log_type="错误"):
            status_item.setText(status)
            # 存储错误信息
            status_item.setData(QtCore.Qt.UserRole, f"{host} {username} {text}")
            # 存储错误次数
            count = int(status_item.data(QtCore.Qt.UserRole + 1))
            if log_type == "错误":
                if count <= 10:
                    count += 1
            elif status in ["运行中", "登录成功"]:
                count = 0

            status_item.setData(QtCore.Qt.UserRole + 1, count)
            print(host, username, text, log_type=log_type)

        try:
            if ax_rdp.property("Connected") == 1:
                log("运行中", status="运行中", log_type="信息")
                return

            # ##################################绑定事件################################### #
            # 绑定方式
            def bound(pyqt_bound_signal, func):
                try:
                    # 断开信号槽函数连接
                    pyqt_bound_signal.disconnect()
                except Exception as p:
                    _ = p
                finally:
                    pyqt_bound_signal.connect(func)

            # 断开连接远程桌面会话
            def on_disconnected(disc_reason=None):
                # 人工操作【断开连接】
                if status_item.text() == "断开连接":
                    log("人工操作【断开连接】", status="断开连接", log_type="警告")
                    return
                error_info = DISC_REASON.get(str(disc_reason), "未知错误")
                count = int(status_item.data(QtCore.Qt.UserRole + 1)) + 1

                if count >= 10:
                    error_info = "连续登录多次错误，请人工排查原因！"

                log(error_info, status="连接断开", log_type="错误")

                if count < 10:
                    # 重新连接
                    print(host, username, "正在重新连接", log_type="警告")
                    # 密码
                    advanced_settings7.setProperty("ClearTextPassword", password)
                    ax_rdp.dynamicCall("Connect()")

            bound(ax_rdp.OnDisconnected, on_disconnected)

            # 开始连接远程桌面会话
            def on_connecting():
                log("连接中", status="连接中", log_type="信息")

            bound(ax_rdp.OnConnecting, on_connecting)

            # 成功登录到远程桌面会话
            def on_login_complete():
                log("登录成功", status="登录成功", log_type="信息")

            bound(ax_rdp.OnLoginComplete, on_login_complete)

            # 在发生登录错误或其他登录
            def on_logon_error(i_error=None):
                error_info = I_ERROR.get(str(i_error), "未知错误")
                log(error_info)

            bound(ax_rdp.OnLogonError, on_logon_error)

            # 在 ActiveX 控件显示身份验证对话框之前调用
            def on_authentication_warning_displayed():
                error_info = "在 ActiveX 控件显示身份验证对话框之前调用(例如，证书错误对话框)"
                log(error_info)
                return

            bound(ax_rdp.OnAuthenticationWarningDisplayed, on_authentication_warning_displayed)

            # 遇到严重错误
            def on_fatal_error(error_code=None):
                error_info = ERROR_CODE.get(str(error_code), "未知错误")
                log(error_info)

            bound(ax_rdp.OnFatalError, on_fatal_error)

            # 遇到非致命错误条件
            def on_warning(warning_code=None):
                error_info = "未知错误"
                if warning_code == 1:
                    error_info = "位图缓存已损坏"
                log(error_info)

            bound(ax_rdp.OnWarning, on_warning)

            # ##################################登录操作################################### #
            ax_rdp.setControl("{8B918B82-7985-4C24-89DF-C33AD2BBFBCD}")
            # 一般属性
            ax_rdp.setProperty("Server", host)
            ax_rdp.setProperty("UserName", username)
            ax_rdp.setProperty("Domain", domain)

            # 全屏显示
            ax_rdp.setProperty("FullScreen", False)

            # 高级属性
            advanced_settings7 = ax_rdp.querySubObject("AdvancedSettings7")

            # 密码
            advanced_settings7.setProperty("ClearTextPassword", password)

            try:
                ax_rdp.dynamicCall("Refresh()")
            except Exception as e:
                print(e)

            ax_rdp.dynamicCall("Connect()")
            ax_rdp.show()

        except Exception as ex:
            log(str(ex))

    # ********定时检测远程情况********* #
    def check_remote(self):
        # 缓存错误信息
        messages = []

        try:
            # 定时器：重新开始计时
            timer_per = int(self.edit_per.text()) * 60
            self.timer.start(1000 * timer_per)

            print("监测开始", log_type="调试")

            # 循环表格
            row_count = self.table_desk.rowCount()
            for row in range(row_count):

                # 获取当前行的会话信息
                host = self.table_desk.item(row, 0).text()
                username = self.table_desk.item(row, 1).text()
                password = self.table_desk.item(row, 2).data(QtCore.Qt.UserRole)
                domain = self.table_desk.item(row, 3).text()
                status_item = self.table_desk.item(row, 4)
                # 状态
                status_text = status_item.text()
                ax_rdp = self.table_desk.cellWidget(row, 5)

                if status_text == "未连接":
                    print("监测中", host, username, "未连接，无需监测会话状态")
                    continue

                if status_text == "断开连接":
                    print("监测中", host, username, "人工操作【断开连接】，请人工重新连接")
                    messages.append(f"{host} {username} 人工操作【断开连接】，请人工重新连接")
                    continue

                # 当前会话的连接状态。
                connected_status = ax_rdp.property("Connected")

                if connected_status == 1:
                    print("监测中", host, username, "会话连接正常")

                elif connected_status == 2:
                    print("监测中", host, username, "连接中，无需监测会话状态")

                elif connected_status == 0:
                    # 错误信息
                    error_info = status_item.data(QtCore.Qt.UserRole)
                    if error_info:
                        messages.append(error_info)

                    # 会话异常，重新连接
                    print("监测中", host, username, "远程会话未连接，尝试连接中", log_type="警告")
                    self.connect_event(host, username, password, domain, ax_rdp, status_item)

        except Exception as ex:
            print("定时器监测", str(ex), log_type="错误")
            messages.append(f"定时器异常{str(ex)}")
        finally:
            if messages:
                webhook = self.edit_webhook.text()
                secret = self.edit_secret.text()
                self.send_dingtalk_robot_info(webhook, secret, ding_talk_text="\n".join(messages))
            print("监测结束", log_type="调试")


if __name__ == "__main__":

    app = QApplication(sys.argv)

    # 应用程序是否启用
    server_name = f"QtRemote_{getpass.getuser()}"
    local_socket = QLocalSocket()
    local_socket.connectToServer(server_name)
    if local_socket.waitForConnected(1000):
        print("QtRemote 已启动")
        sys.exit(0)
    else:
        local_server = QLocalServer()
        local_server.listen(server_name)
        print("QtRemote 打开中")
    window = Main()
    window.show()
    sys.exit(app.exec_())
