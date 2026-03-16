1. Python 第三方库（通过 pip 安装）
Flask==3.0.0
Flask-SQLAlchemy==3.1.1
Flask-Migrate==4.0.5
pandas==2.1.3
openpyxl==3.1.2
selenium==4.15.2
requests==2.31.0
urllib3==2.1.0
pdf2image==1.16.3
Pillow==10.1.0
Werkzeug==3.0.1
SQLAlchemy==2.0.23

2. 系统级软件（需手动下载安装）
a) FFmpeg（语音功能必需）
下载地址：https://ffmpeg.org/download.html
安装后需要配置环境变量，或者在系统配置表中设置 FFmpeg 路径
代码中默认路径：D:\ffmpeg\bin\ffmpeg.exe

b) Poppler（pdf2image 的依赖，OCR 功能必需）
Windows 版本下载地址：http://blog.alivate.com.au/poppler-windows/
下载后解压，将 bin 目录添加到系统 PATH 环境变量

c) Microsoft Edge WebDriver（如果需要自动爬取功能）
下载与 Edge 浏览器版本匹配的 WebDriver
下载地址：https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/
