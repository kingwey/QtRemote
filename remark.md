# 一、ui转py文件
python -m PyQt5.uic.pyuic E:\QtRemote\ui\main.ui -o E:\QtRemote\ui\main.py

# 二、pip
- PyQt5: PyQt5
- clr: pythonnet
- ~~pythoncom：pywin32~~
- ~~pythonping: pythonping~~
- psutil: psutil
- requests: requests
- rsa: rsa
- Crypto: pycryptodome

# 三、root key
## public_root
("LS0tLS1CRUdJTiBSU0EgUFVCTElDIEtFWS0tLS0tCk1FZ0NRUUNnUVJQZHFaZE1HeGs0dzRLYkIy"
"cmtTNWJKenE2SVZOai9BcXhYdlhIR1RVZjNvbGdNN3o3UURNNDMKRTRhQjFEd3BDV0ZEajRia2o5"
"YlBpdmE2SHRrVkFnTUJBQUU9Ci0tLS0tRU5EIFJTQSBQVUJMSUMgS0VZLS0tLS0K")

## private_root
("LS0tLS1CRUdJTiBSU0EgUFJJVkFURSBLRVktLS0tLQpNSUlCUEFJQkFBSkJBS0JCRTkycGwwd2JH"
"VGpEZ3BzSGF1Ukxsc25Pcm9oVTJQOENyRmU5Y2NaTlIvZWlXQXp2ClB0QU16amNUaG9IVVBDa0pZ"
"VU9QaHVTUDFzK0s5cm9lMlJVQ0F3RUFBUUpBTTkyOEswTEhTQWVCTzBEejFXOHEKSm1kY2owWklj"
"TEZkWmZPY2llMHpXbTY4cUpsWE5BQ1E1ems3bUtWcnA3LzhPRk5TY3RoVDdrNmYrZDRpUnJvMgpa"
"UUlqQUxQZkFPNjNtb3oyMys1dnY4NEVsVFNVTmczSkRyaXNxWi9pancxRi9kK1hUWThDSHdEa0ZK"
"WlFKckNKCjJDTDQ1b3ZJOWlLWVNSbzE4NU9hVVBBRndIZk1KUnNDSXdDTUdCQzkzVHIrdC9uSjJE"
"Zm4yaUhzQmRQa0FNajYKaFdESUtzbUlhUTlHNnExNUFoNXJ6ak5TUlVkU2tGL1BhQ0dRWm83cGpq"
"d2VYamhaUzRKNEpZWTZieHNDSW1kZwp5Y3ZPRjA0d2VUQlAvQlBTaTJ6RUNaWnJ1aGdFUFVISyt2"
"a1RINlI1NmFNPQotLS0tLUVORCBSU0EgUFJJVkFURSBLRVktLS0tLQo=")

# 四、pyd编译

## 安装第三方包
pip install cython -i https://pypi.tuna.tsinghua.edu.cn/simple

## 编写setup.py文件
from distutils.core import setup
from Cython.Build import cythonize
setup(ext_modules=cythonize(["E:\QtRemote\common.py"]))

## 命令行：执行编译
python setup.py build_ext --inplace

# 五、自动生成pyi文件

## 安装第三方包
pip install mypy -i https://pypi.tuna.tsinghua.edu.cn/simple
## 命令行：生成pyi文件（待生成pyi的py文件）
stubgen E:\QtRemote\clr\common.py

# 六、打包exe

## 安装第三方包
pip install pyinstaller==5.4.0 -i https://pypi.tuna.tsinghua.edu.cn/simple
## 命令行：切换至项目目录
cd /d E:\python-code\QtRemote
## 命令行：生成exe
pyinstaller -w -i "./img/logo.ico" -n QtRemote –-distpath "./" --uac-admin "main.py" 
--add-data "domain.pem;.\\\\." 
--add-data "ui;ui" 
--add-data "css;css" 
--add-data "img;img" 
--add-data "core;core" 
--add-data "E:\python-code\QtRemote\venv\Lib\site-packages\rsa;rsa" 
--hidden-import hmac 
--hidden-import json 
--hidden-import psutil 
--hidden-import requests 
--hidden-import Crypto.Cipher.PKCS1_v1_5 
--hidden-import Crypto.PublicKey.RSA 
--hidden-import win32process 
--hidden-import win32gui 
-y

# 七、制作安装包

## 【开始】菜单
C:\ProgramData\Microsoft\Windows\Start Menu\Programs

