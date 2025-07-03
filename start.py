import ctypes
import os
import subprocess
import sys
from datetime import datetime
import tempfile
import time
import re


def del_file(lee_path, before_days=30, ext_name=".*", time_type="创建日期", is_recursion=False):
    """
    * 方法:删除N天前创建的文件
    * 参数:
            * @lee_path: 待删除文件夹
            * @before_days: N天前
            * @ext_name: 待删除文件的后缀名
            * @time_type: 日期类型
            * @is_recursion: 是否递归文件夹

    * 返回:
    """

    if os.path.exists(lee_path):
        for ifile in os.listdir(lee_path):
            try:
                full_file = os.path.join(lee_path, ifile)

                if os.path.isdir(full_file):
                    if is_recursion is True:
                        del_file(full_file, before_days, ext_name, time_type, is_recursion)
                        if not os.listdir(full_file):
                            os.removedirs(full_file)
                    continue

                # 获取文件相关日期
                if time_type == "创建日期":
                    file_date = os.path.getctime(full_file)
                elif time_type == "修改日期":
                    file_date = os.path.getmtime(full_file)
                else:
                    # 访问日期
                    file_date = os.path.getatime(full_file)

                if file_date < time.time() - 60 * 60 * 24 * int(before_days):
                    if ".*" in ext_name:
                        os.remove(full_file)
                    elif os.path.splitext(full_file)[-1].lower() in ext_name:
                        os.remove(full_file)

            except Exception as ex:
                print(ex.args[0])

    else:
        print(lee_path + "：路径不存在")


def print(*args, sep=" ", end="\n", file=sys.stdout, log_type="信息"):
    """
    * 方法:打印日志
    * 参数:
            * @*args: 日志信息
            * @sep: 分隔符
            * @end: 结束符
            * @file: 日志打印到
            * @log_type: 日志类型（调试，信息，警告，错误）

    * 返回:
    """
    try:
        # 日志时间
        local_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        # 日志类型
        log_type = log_type if log_type else "信息"

        # 日志信息
        args = [str(value) for value in args]
        text = sep.join(args) + end
        # 刷新日志记录
        r = "\r" if text.startswith("\r") else ""
        # 格式化
        text = f"{r}{local_time}[{log_type}]:" + text.strip("\r")

        # 日志文件夹
        log_path = os.path.join(tempfile.gettempdir(), "leqee_log")
        if not os.path.exists(log_path):
            os.mkdir(log_path)

        # 日志文件路径
        log_file = os.path.join(log_path, local_time.split(' ')[0] + ".log")

        # 删除历史日志文件
        if not os.path.exists(log_file):
            del_file(log_path)

        # 打印日志至文件
        with open(log_file, mode="a", encoding="utf-8", errors="ignore") as f:
            f.write(f"{text.strip()}\n")

        try:
            # 打印日志至控制台
            file.write(text)
        except:
            pass

    except Exception as ex:
        print(str(ex))


def latest_version():
    """获取最新版本"""
    # 初试版本
    tmp_version_value = None
    # 初试app
    latest_version_exe = sys.argv[0]

    # 根目录
    root_path = os.path.dirname(latest_version_exe)
    print("根目录", root_path)

    # 根应用名称
    root_name = os.path.splitext(os.path.basename(latest_version_exe))[0]
    print("根应用名称", root_name)

    for folder in os.listdir(root_path):

        if os.path.isdir(os.path.join(root_path, folder)) and folder.startswith(root_name):

            # 正则匹配数字
            version_value = re.findall(r"\d+", folder)

            if tmp_version_value is None:
                # 初试化版本号
                tmp_version_value = [0 for i in version_value]

            # 判断最新版本
            if (int(tmp_version_value[0]) <= int(version_value[0]) and
                    int(tmp_version_value[1]) <= int(version_value[1]) and
                    int(tmp_version_value[2]) <= int(version_value[2])):

                latest_version_exe = os.path.join(root_path, folder, f"{root_name}.exe")
                tmp_version_value = version_value

    return latest_version_exe


def admin_runas():
    """管理员权限启动"""
    # 判断是否是管理员权限
    if ctypes.windll.shell32.IsUserAnAdmin():
        latest_version_exe = latest_version()
        # 防止死循环
        if latest_version_exe != sys.argv[0]:
            print("最新版", latest_version_exe)
            subprocess.Popen(latest_version_exe)
        else:
            print("搜索不到最新版exe文件，无法启动！", log_type="错误")
    else:
        print("非管理员权限", log_type="错误")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 0)


# 按装订区域中的绿色按钮以运行脚本。
if __name__ == '__main__':
    admin_runas()

