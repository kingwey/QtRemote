import sys
import os
import shutil
import oss2
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QFileDialog, QProgressBar
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QThread, pyqtSignal


class DownloadThread(QThread):
    progress_signal = pyqtSignal(int)
    finish_signal = pyqtSignal(str)

    def __init__(self, access_key_id, access_key_secret, endpoint, bucket_name, object_name, local_path):
        super().__init__()
        self.access_key_id = access_key_id
        self.access_key_secret = access_key_secret
        self.endpoint = endpoint
        self.bucket_name = bucket_name
        self.object_name = object_name
        self.local_path = local_path

    def run(self):
        try:
            auth = oss2.Auth(self.access_key_id, self.access_key_secret)
            bucket = oss2.Bucket(auth, self.endpoint, self.bucket_name)

            def progress_callback(consumed_bytes, total_bytes):
                if total_bytes:
                    progress = int(100 * (consumed_bytes / total_bytes))
                    self.progress_signal.emit(progress)

            result = bucket.get_object_to_file(self.object_name, self.local_path, progress_callback=progress_callback)
            if result.status == 200:
                self.finish_signal.emit("下载完成！")
            else:
                self.finish_signal.emit(f"下载失败，状态码: {result.status}")
        except Exception as e:
            self.finish_signal.emit(f"下载出错: {str(e)}")


class InstallerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.install_path = ""
        self.download_thread = None

    def initUI(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f0f0;
            }
            QLabel {
                font-size: 14px;
                color: #333;
                padding: 10px;
            }
            QPushButton {
                background-color: #007BFF;
                color: white;
                font-size: 14px;
                padding: 8px 16px;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #0056b3;
            }
            QProgressBar {
                border: 1px solid #ccc;
                border-radius: 4px;
                background-color: #f5f5f5;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #28a745;
                border-radius: 3px;
            }
        """)

        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(14)

        self.info_label = QLabel("欢迎使用安装程序，请选择安装路径。")
        self.info_label.setFont(font)
        layout.addWidget(self.info_label)

        self.select_path_button = QPushButton("选择安装路径")
        self.select_path_button.setFont(font)
        self.select_path_button.clicked.connect(self.select_install_path)
        layout.addWidget(self.select_path_button)

        self.download_button = QPushButton("从 OSS 下载并安装")
        self.download_button.setFont(font)
        self.download_button.clicked.connect(self.start_download)
        self.download_button.setEnabled(False)
        layout.addWidget(self.download_button)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setFont(font)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)

        self.setWindowTitle('安装程序')
        self.setGeometry(300, 300, 400, 280)
        self.show()

    def select_install_path(self):
        path = QFileDialog.getExistingDirectory(self, "选择安装路径")
        if path:
            self.install_path = path
            self.info_label.setText(f"已选择安装路径: {self.install_path}")
            self.download_button.setEnabled(True)

    def start_download(self):
        if self.install_path:
            self.info_label.setText("开始下载...")
            # 替换为你的 OSS 信息
            access_key_id = 'your_access_key_id'
            access_key_secret = 'your_access_key_secret'
            endpoint = 'your_endpoint'
            bucket_name = 'your_bucket_name'
            object_name = 'your_object_name'
            local_path = os.path.join(self.install_path, object_name)

            self.download_thread = DownloadThread(access_key_id, access_key_secret, endpoint, bucket_name, object_name,
                                                  local_path)
            self.download_thread.progress_signal.connect(self.update_progress)
            self.download_thread.finish_signal.connect(self.download_finished)
            self.download_thread.start()

    def update_progress(self, progress):
        self.progress_bar.setValue(progress)

    def download_finished(self, message):
        self.info_label.setText(message)
        if message.startswith("下载完成！"):
            # 这里可以添加安装逻辑，例如解压文件等
            pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    installer = InstallerApp()
    sys.exit(app.exec_())
