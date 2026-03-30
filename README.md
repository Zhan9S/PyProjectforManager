基于 Flask 框架开发的高校教学成果信息化管理平台，集成 OCR 智能识别、语音导出、团队协同、数据统计等功能，为高校教师和管理者提供一站式成果管理解决方案。
负责系统架构设计、数据库建模、核心功能开发及部署运维全流程工作。采用 Flask + SQLAlchemy 技术栈，使用 SQLite 作为数据库，集成 Pandas 进行数据处理，pdf2image 和 Pillow 处理 PDF 和图片，FFmpeg 进行音频处理，Selenium 实现浏览器自动化数据采集。
实现 OCR 智能识别模块，通过正则表达式和关键词匹配算法自动识别上传材料中的成果类型并提取关键信息，支持期刊论文、发明专利、教材、专著、软著、教学成果获奖等十余种成果类型的自动化录入。开发语音导出功能，将成果数据转换为音频格式用于汇报展示场景。设计三级权限体系（普通教师、团队负责人、管理员），实现用户管理、团队管理、字典管理等完整功能模块。构建数据统计仪表盘，多维度可视化展示个人及团队成果数据。
系统已稳定运行，支持用户上传各类教学成果材料，通过 OCR 识别大幅减少手动录入工作量，为高校教师职称评审、年度考核、成果申报等工作提供高效的数据支撑。

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
FFmpeg（语音功能必需）
下载地址：https://ffmpeg.org/download.html
安装后需要配置环境变量，或者在系统配置表中设置 FFmpeg 路径
代码中默认路径：D:\ffmpeg\bin\ffmpeg.exe

Poppler（pdf2image 的依赖，OCR 功能必需）
Windows 版本下载地址：http://blog.alivate.com.au/poppler-windows/
下载后解压，将 bin 目录添加到系统 PATH 环境变量

Microsoft Edge WebDriver（如果需要自动爬取功能）
下载与 Edge 浏览器版本匹配的 WebDriver
下载地址：https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/


