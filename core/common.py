import base64
import copy
import csv
import getpass
import hashlib
import hmac
import json
import os
import socket
import sys
import time
import urllib.parse
from datetime import datetime

import psutil
import requests
import rsa
from Crypto.Cipher import PKCS1_v1_5 as Cipher_pkcs1_v1_5
from Crypto.PublicKey import RSA


class Crypt:

    def __init__(self):
        # ../img/logo.ico
        self.load_path = os.path.dirname(__file__)
        self.key_file = os.path.join(self.load_path, r"config.key")
        # 需要人工提供
        self.public_root = ""
        self.private_root = ("LS0tLS1CRUdJTiBSU0EgUFJJVkFURSBLRVktLS0tLQpNSUlCUEFJQkFBSkJBS0JCRTkycGwwd2JH"
                             "VGpEZ3BzSGF1Ukxsc25Pcm9oVTJQOENyRmU5Y2NaTlIvZWlXQXp2ClB0QU16amNUaG9IVVBDa0pZ"
                             "VU9QaHVTUDFzK0s5cm9lMlJVQ0F3RUFBUUpBTTkyOEswTEhTQWVCTzBEejFXOHEKSm1kY2owWklj"
                             "TEZkWmZPY2llMHpXbTY4cUpsWE5BQ1E1ems3bUtWcnA3LzhPRk5TY3RoVDdrNmYrZDRpUnJvMgpa"
                             "UUlqQUxQZkFPNjNtb3oyMys1dnY4NEVsVFNVTmczSkRyaXNxWi9pancxRi9kK1hUWThDSHdEa0ZK"
                             "WlFKckNKCjJDTDQ1b3ZJOWlLWVNSbzE4NU9hVVBBRndIZk1KUnNDSXdDTUdCQzkzVHIrdC9uSjJE"
                             "Zm4yaUhzQmRQa0FNajYKaFdESUtzbUlhUTlHNnExNUFoNXJ6ak5TUlVkU2tGL1BhQ0dRWm83cGpq"
                             "d2VYamhaUzRKNEpZWTZieHNDSW1kZwp5Y3ZPRjA0d2VUQlAvQlBTaTJ6RUNaWnJ1aGdFUFVISyt2"
                             "a1RINlI1NmFNPQotLS0tLUVORCBSU0EgUFJJVkFURSBLRVktLS0tLQo=")

    def load_key(self):
        """对key进行解密处理"""
        key_list = [
            "RcfZKiTrTI7NiXgNhPUX5hD4uGrMmGiXR+OVL03efPQAFRgYpyoSk+DB0c+RJs9fl3IQI/oLk2il5t+aJKvACw==",
            "Uhgins02l7VFXi9n3q+ltUN7U9zFIJjuSqJpu849E72VGfg3l6kuoX/jWODMN7/CUMeh1JIRgUQRlJFNfImdtA==",
            "ExV4ztZbBBnO4/BJCeofpwbzneHpvAW1IUaYcT2BIZGjq0I8seMOWW4MW+CMjHRzKtlKh6hkoKf4l+PdWQv51w==",
            "Y5Z3wr+O0YfE+bpYgSz2TUN2FqlBGDYftSKUO2KwYs3oHRbT+f83+nLGeRJyRwGU3h9GVLAYprQNSGO/wPVnNQ==",
            "VwfboXwiDKaE7kho7WEPk56k4QYTOfLfBVA4hH7gKzQ4LJQUkXUMXi+rdc2/3/NR3zv6QxtW5903adlGxVK9Cw==",
            "ToTJeVocOiPJzedTi/HwjsBrPE0YW6N8MLjPsNohtuCJwC/ZzTr04KBt188KCbDTgXq+Y1kdFyir3lZ1rN7XBg==",
            "WjI3VBstoIO2g7Ayl5CDzC4JHx2OhySIRabAjq2ObUe/rIUu6fK1BthxKytlnKFktHXEAsFlX+sHniJS7KD4iA==",
            "LW/KToazXNRSXFAzLhGLvqndAaGSifRK82+3hu0oY9FASTMxCw4pMw+uhzCDULiJwCgdQpC5UgX7d5Po+m7jFw==",
            "PDZRmv58vB5p+2uPRWTV4s85Ywf/3kpJS8S4qgfgb1VIpxjeZduxBt0qwYp44CwgE6NB0ZiD5V2xJw/KEI0ZZA==",
            "Hf08OZoSfLxumyoCQIeq4YDflRAtR5+HlPyM8Q3KCj6qTTZ5jxNPOc/c5+OeGpqazuNCXg13OgcRL0xtdfxh+g==",
            "MOZB/y7Ms9BzJ3j/Xh64PUEzXDb2zVxxNzU15yLxGAwtpy78p0NR+yd3GEeKzEe7Lo5aas9aI3dm1aejfst6cA==",
            "cuI8JATcIEEmXzBXOAcBQV1qICa/wCQgGJ0plWj6jwzkjSdGREaeaGUHefv/+km1S8O9hXeJ/oiC4BV+woHyMw==",
            "PGlRN4KXVnV6SKX7IlnmOG8Rlom63QYfYi5xfR6EaMM9Y4pwmvwvAo6YAsdSyq6Ka/6zVOZJRe2IdV+CvuMaJw==",
            "jYceoijRrobXBlBb48dDqNEWuyFw/DYyLEpQdfm36l2Gz5SnfxhRaEQkR/XibrSE7j7Nhcia0Pw3j1l5SqT7Bg==",
            "TdGaARDUKDj6gt7YGqBrllKCPGiKdPsIw1wsIhZQfWxbmijjYOY6FNTHd3vUuMYGJ0FoGaQl1xQbOY7oi1scdQ==",
            "AfPo9sMgejmJJzViZJ1UWIbwqVrWjZiSnmu8qD1Aaxt552nYuksaq8EISD5Q0vxNYxIDN7fDayDEnqy2Q6mMGA==",
            "YanKQN79+hlpgY9JbYoAbx7I8OFSasBMCV8FvBe7GWe0+w5VhK4HxX3tl9lDg2zWEmQMCjRKaIMT9F91Mk7IYw==",
            "Tfm3HFCWdxIPNTmHboqzIk+t9e8Ct7m4dKe0dKW/nEhOrOl/IOsFKBLuVrGzLAGQlqnxSCOIt1Rm7shBqIT17Q==",
            "g2aBwY8I4UrYIvpbhlF/1rbnWT65dOxpsvgoLPaXJFGCTEBFnfGD/c/nhpYxfwIOE9SVzQNhyWLpWZU+1xCCEA=="
        ]
        key_text = ""
        for key_item in key_list:
            key_text += self.rsa_decrypt(self.private_root, key_item.strip())
        # 返回：公钥，私钥
        return json.loads(key_text)

    def dump_key(self, public_key, private_key):
        """对key进行加密处理"""
        key_json = {
            "public_key": public_key,
            "private_key": private_key
        }
        key_text = json.dumps(key_json)
        split_length = 50

        with open(self.key_file, mode="w", encoding="utf-8", errors="ignore") as f:
            for i in range(0, len(key_text), split_length):
                item_key = key_text[i:i + split_length]
                f.writelines(self.rsa_encrypt(self.public_root, item_key) + "\n")

    @classmethod
    def create_rsa_key(cls):
        """
        方法:生成ras公钥私钥对

        返回:
            - public_key: 公钥
            - private_key: 私钥
        """
        # 生成RSA公钥和秘钥,经过base64转码
        (public_key, private_key) = rsa.newkeys(512)

        # 经过base64编码
        public_key = base64.encodebytes(public_key.save_pkcs1())
        private_key = base64.encodebytes(private_key.save_pkcs1())

        return public_key, private_key

    @classmethod
    def rsa_encrypt(cls, public_key, plain_text):
        """
        方法:使用rsa公钥加密

        参数:
            - @public_key: 公钥
            - @plain_text: 待加密的明文

        返回:
            - cipher_text:加密后的密文
        """

        cipher = Cipher_pkcs1_v1_5.new(RSA.importKey(base64.b64decode(public_key)))
        cipher_text = base64.b64encode(cipher.encrypt(plain_text.encode())).decode()

        return cipher_text

    @classmethod
    def rsa_decrypt(cls, private_key, cipher_text):
        """
        方法:使用rsa私钥解密

        参数:
            - @private_key: 私钥
            - @cipher_text: 待解密的密文

        返回:
            - plain_text: 解密后的明文
        """
        try:
            cipher = Cipher_pkcs1_v1_5.new(RSA.import_key(base64.b64decode(private_key)))
            plain_text = cipher.decrypt(base64.b64decode(cipher_text), 'ERROR').decode('utf-8')
            return plain_text
        except Exception as ex:
            _ = ex
            return False


class Common:
    """通用类"""
    def __init__(self):

        # 本机主机
        # 本机当前会话用户名
        self.username = getpass.getuser()
        # 本机主机名
        self.hostname = socket.gethostname()
        # 本机物理ip地址
        self.host = socket.gethostbyname(self.hostname)

        # 进程
        # 应用程序进程对象
        self.process = psutil.Process()
        # 应用程序进程pid
        self.process_pid = self.process.pid
        # 测试进程名
        self.test_process_name = "LocalTest.exe"
        # 应用程序进程名
        self.process_name = self.test_process_name if self.process.name() == "python.exe" else self.process.name()

        # 应用属性
        # 应用程序窗口标题
        self.window_title = f"QtRemote_{self.host}_{self.username}"

        # # 应用程序是否启用
        # self.is_enable = True if win32gui.FindWindow(None, self.window_title) else False

        # 应用配置
        # 配置文件路径
        self.preferences_file = os.path.join(self.user_file, "preferences")
        # 配置文件全部信息
        self.preferences_info = {}

    @property
    def local_date(self):
        """当前日期（%Y-%m-%d）"""
        return datetime.now().strftime('%Y-%m-%d')

    @property
    def local_time(self):
        """当前时间（%Y-%m-%d %H:%M:%S）"""
        return datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    @property
    def user_file(self):
        """当前进程用户文件夹"""
        _ = os.path.join(os.environ["USERPROFILE"], "AppData\\Local", self.process_name.replace(".exe", ""))
        # 文件夹不存在则创建
        if not os.path.exists(_):
            os.makedirs(_)
        return _

    @property
    def log_file(self):
        """日志文件"""
        _ = os.path.join(self.user_file, "log")
        if not os.path.exists(_):
            os.makedirs(_)
        # 清理日志文件
        self.del_file(_)
        return os.path.join(_, self.local_date + ".log")

    def print(self, *args, sep=" ", end="\n", file=sys.stdout, log_type="信息"):
        """
        方法:打印日志

        参数:
            - @*args: 日志内容
            - @sep: 分割符
            - @end: 结束符
            - @file: 打印至。。。
            - @log_type: 日志类型（调试，信息，警告，错误）

        返回:
            - 无
        """

        try:
            # 日志类型
            log_type = log_type if log_type else "信息"

            # 日志信息
            args = [str(value) for value in args]
            text = sep.join(args) + end

            # 日志格式化
            r = "\r" if text.startswith("\r") else ""
            text = f"{r}{self.local_time}[{log_type}]:{text}"

            # 打印日志至文件
            with open(self.log_file, mode="a", encoding="utf-8", errors="ignore") as f:
                f.write(text.strip() + "\n")

            # 打印日志至控制台
            try:
                file.write(text)
            except Exception as ex:
                _ = ex

        except Exception as ex:
            _ = ex

    def del_file(self, lee_path, before_days=30, ext_name=".*", time_type="创建日期", is_recursion=False):
        """
        方法:删除N天前创建的文件

        参数:
            - @lee_path: 待删除文件夹
            - @before_days: N天前
            - @ext_name: 待删除文件的后缀名
            - @time_type: 日期类型（创建日期，修改日期，最后一次访问日期）
            - @is_recursion: 是否递归删除文件夹

        返回:
            - 无
        """
        if not os.path.exists(lee_path):
            print(lee_path + "：路径不存在")
            return

        for ifile in os.listdir(lee_path):
            try:
                full_file = os.path.join(lee_path, ifile)

                # 文件夹处理
                if os.path.isdir(full_file):
                    if is_recursion is True:
                        # 递归删除文件夹
                        self.del_file(full_file, before_days, ext_name, time_type, is_recursion)

                        # 删除空文件夹
                        if not os.listdir(full_file):
                            os.removedirs(full_file)
                    continue

                # 获取文件相关日期
                if time_type == "创建日期":
                    file_date = os.path.getctime(full_file)
                elif time_type == "修改日期":
                    file_date = os.path.getmtime(full_file)
                else:   # 访问日期
                    file_date = os.path.getatime(full_file)

                if file_date < time.time() - 60 * 60 * 24 * int(before_days):
                    if ".*" in ext_name:
                        os.remove(full_file)
                    elif os.path.splitext(full_file)[-1].lower() in ext_name:
                        os.remove(full_file)

            except Exception as ex:
                self.print(ex.args[0])

    def send_dingtalk_robot_info(self, webhook, secret, ding_talk_text=""):
        """
        方法:钉钉机器人发送群消息（远程桌面异常通知）

        参数:
            - webhook: Webhook地址
            - secret: 秘钥
            - ding_talk_text: 消息内容

        返回:
            * 无

        帮助:
            [官方文档](https://open.dingtalk.com/document/orgapp/custom-robots-send-group-messages)
        """
        try:
            # 签名
            timestamp = str(round(time.time() * 1000))
            secret_enc = secret.encode('utf-8')
            string_to_sign = f'{timestamp}\n{secret}'
            string_to_sign_enc = string_to_sign.encode('utf-8')
            hmac_code = hmac.new(secret_enc, string_to_sign_enc, digestmod=hashlib.sha256).digest()
            sign = urllib.parse.quote_plus(base64.b64encode(hmac_code))

            # 消息内容格式化
            text_content = f"【远程桌面异常通知】{self.local_time}\n主机IP：{self.host} {self.username}\n远程异常：\n{ding_talk_text}"

            # 请求连接
            url = webhook + '&timestamp={}&sign={}'.format(timestamp, sign)
            headers = {'Content-Type': 'application/json'}
            ding_talk = {
                "at": {
                    "atMobiles": [],
                    "atUserIds": [],
                    "isAtAll": True
                },
                "text": {
                    "content": text_content
                },
                "msgtype": "text"
            }
            # 测试
            if self.process_name == self.test_process_name:
                print("钉钉消息\n" + text_content)
                return

            # 发送请求
            print("监测异常", "推送钉钉消息")
            requests.post(url=url, headers=headers, json=ding_talk)
        except Exception as ex:
            _ = ex

    def read_preferences_info(self):
        """读取配置文件"""
        default_preferences = {
            "notice": {
                "webhook": "https://oapi.dingtalk.com/robot/send?access_token=c3f39994904245f2d73847648aa17a8c4404654da1c2dd1b21ee52f2316fa55f",
                "secret": "SEC1187c5c3e7a313a63169265a1e7c6b14bcdd93739dad1c6cf64dfa000be31697",
                "per": "30"
            }
        }
        try:
            if os.path.exists(self.preferences_file):
                with open(self.preferences_file, mode="r", encoding="utf-8") as f:
                    json_preferences = json.loads(f.read())
                    if json_preferences:
                        return json_preferences
                    else:
                        return default_preferences
            else:
                return default_preferences
        except:
            return default_preferences

    def save_preferences_info(self):
        """保存配置文件"""
        with open(self.preferences_file, mode="w", encoding="utf-8") as f:
            f.write(json.dumps(self.preferences_info))

    def get_preferences_keys(self, keys=[]):
        """获取指定键的配置信息"""
        try:
            self.preferences_info = self.read_preferences_info()
            do = "self.preferences_info"
            for key in keys:
                do += f".get('{key}')"
            return eval(do)
        except Exception as ex:
            print(ex)

    def set_preferences_keys(self, keys=[], value=None):
        """设置指定键的配置信息"""
        try:
            local_vars = {"value": value, "self": self}
            do = "self.preferences_info"
            for key in keys:
                do += f"['{key}']"
            do += f"=value"
            exec(do, globals(), local_vars)
            self.save_preferences_info()
        except Exception as ex:
            print(ex)


class HostRemote(Common, Crypt):
    """继承Common类"""

    def __init__(self):
        super(Common, self).__init__()
        super().__init__()

        self.header = ['地址', '用户名', '密码', '域名']
        self.host_info = []
        self.host_file = os.path.join(self.user_file, "hostsession")
        # 公钥、私钥
        key_json = self.load_key()
        self.public_key = key_json["public_key"]
        self.private_key = key_json["private_key"]
        # 初始化host session文件
        self.load()
        self.dump()

    def refresh(func):
        """刷新数据"""

        def inner(self, *args, **kwargs):
            result = func(self, *args, **kwargs)
            self.dump()
            self.load()
            return result

        return inner

    @refresh
    def add(self, row):
        """增加一行"""
        if len(row) == len(self.header):
            self.host_info.append(row)
        else:
            print(f"增加信息长度不匹配", log_type="错误")

    @refresh
    def delete(self, row):
        """删除一行"""
        for desk_info in self.host_info:
            if desk_info[0] == row[0] and desk_info[1] == row[1] and desk_info[3] == row[3]:
                self.host_info.remove(desk_info)
                break
        else:
            print(f"索引值大小超出预期", log_type="错误")

    @refresh
    def modify(self, index, row):
        """修改行数据"""
        if len(row) != len(self.header):
            print(f"增加信息长度不匹配", log_type="错误")
            return

        if len(self.host_info) > index >= 0:
            self.host_info[index] = row
        else:
            print(f"索引值大小超出预期", log_type="错误")
        pass

    def select(self, row):
        """查询行数据是否存在"""
        for desk_info in self.host_info:
            if desk_info[0] == row[0] and desk_info[1] == row[1] and desk_info[3] == row[3]:
                return True
        else:
            return False

    def load(self):
        """加载本地数据"""
        try:
            self.host_info = []

            if os.path.exists(self.host_file):
                with open(self.host_file, mode='r', newline='', encoding="UTF-8", errors='ignore') as f:
                    # csv读取器
                    reader = csv.reader(f, delimiter=",")
                    for row in list(reader)[1:]:
                        # 解密
                        password = self.rsa_decrypt(self.private_key, row[2])
                        if password:
                            row[2] = password
                        # 去重判断
                        if row not in self.host_info:
                            self.host_info.append(row)

        except Exception as ex:
            print(f"加载本地数据：{ex}", log_type="错误")

    def dump(self):
        """更新本地数据"""
        try:
            host_info = copy.deepcopy(self.host_info)

            # 加密保存
            for row in host_info:
                # 解密
                if self.rsa_decrypt(self.private_key, row[2]) is False:
                    # 加密
                    row[2] = self.rsa_encrypt(self.public_key, row[2])

            with open(self.host_file, mode='w', newline='', encoding="UTF-8", errors='ignore') as f:
                csvwriter = csv.writer(f)
                csvwriter.writerows([self.header] + host_info)
        except Exception as ex:
            print(f"更新本地数据：{ex}", log_type="错误")


# 打印日志方法
print = Common().print

if __name__ == "__main__":
    # c = Common()
    my_list = [[3, 2], [1, 4], [1, 2], [5, 6]]
    my_list = sorted(my_list, key=lambda x: (x[0], x[1]))
    print(my_list)
