Word to PDF Converter
程序功能
Word to PDF Converter 是一个基于 PyQt6 的图形界面工具，用于将多种文件格式转换为 PDF 文件。支持以下功能：

支持的文件格式：
Microsoft Office 文件：.doc、.docx、.ppt、.pptx、.xls、.xlsx
文本文件：.txt
网页文件：.html、.mhtml
多种转换工具：
使用 LibreOffice 转换多种文件格式（.docx、.pptx、.html 等）。
使用 wkhtmltopdf 转换网页文件（.html 和 .mhtml）。
使用 Microsoft Office 和 WPS Office 的 COM 接口转换 Office 文件（优先 WPS）。
图形界面：
提供简洁的 PyQt6 界面，支持文件选择、转换进度显示。
进度条平滑动画，显示转换进度。
日志记录：
每次运行生成 converter.log 日志文件，记录转换过程和错误信息。
运行环境
操作系统：Windows 10 或更高版本（本程序仅支持 Windows）。
Python 版本：Python 3.8 或更高版本。
磁盘空间：至少 1GB 可用空间（用于安装依赖工具）。
依赖安装
1. Python 依赖
程序需要以下 Python 库，使用阿里云镜像加速安装：

bash

Copy
# 安装 PyQt6（GUI 库）
pip install PyQt6 -i https://mirrors.aliyun.com/pypi/simple/

# 安装 pywin32（用于 COM 接口）
pip install pywin32 -i https://mirrors.aliyun.com/pypi/simple/

# 安装 pdfkit（调用 wkhtmltopdf 转换网页文件）
pip install pdfkit -i https://mirrors.aliyun.com/pypi/simple/
2. 外部依赖
程序依赖以下外部工具，用户需手动安装并配置：

2.1 LibreOffice
用途：转换 .docx、.pptx、.html 等文件。
下载：
访问 LibreOffice 官方网站：https://www.libreoffice.org/download/download/
下载 Windows 版本（64 位或 32 位）。
安装：
双击安装包，按照提示安装（建议保持默认路径：C:\Program Files\LibreOffice）。
配置 PATH：
右键“此电脑” → “属性” → “高级系统设置” → “环境变量”。
在“系统变量”中找到 Path，点击“编辑”。
添加：C:\Program Files\LibreOffice\program。
点击“确定”关闭窗口。
验证：
bash

Copy
soffice --version
应显示版本信息，例如：LibreOffice 7.6.7.2。
2.2 wkhtmltopdf
用途：转换 .html 和 .mhtml 文件。
下载：
访问 wkhtmltopdf 官方网站：https://wkhtmltopdf.org/downloads.html
下载 Windows 版本（例如 wkhtmltox-0.12.6-2.msvc2015-win64.exe）。
安装：
双击安装包，按照提示安装（建议保持默认路径：C:\Program Files\wkhtmltopdf）。
配置 PATH：
右键“此电脑” → “属性” → “高级系统设置” → “环境变量”。
在“系统变量”中找到 Path，点击“编辑”。
添加：C:\Program Files\wkhtmltopdf\bin。
点击“确定”关闭窗口。
验证：
bash

Copy
wkhtmltopdf --version
应显示版本信息，例如：wkhtmltopdf 0.12.6 (with patched qt)。
2.3 Microsoft Office 或 WPS Office
用途：转换 .docx、.pptx、.xlsx 等文件。
Microsoft Office：
如果已有 Office（例如 Office 365、2016、2019），无需额外安装。
否则，请购买并安装 Microsoft Office（需包含 Word、PowerPoint、Excel 组件）。
WPS Office：
下载：https://www.wps.com/download/
安装时确保包含“演示”组件（对应 PowerPoint）。
打包依赖和指令
1. 打包依赖
PyInstaller：用于将 Python 脚本打包为 .exe 文件。
安装 PyInstaller（使用阿里云镜像）：
bash

Copy
pip install pyinstaller -i https://mirrors.aliyun.com/pypi/simple/
2. 打包指令
假设 word_to_pdf_converter_office.py 和图标文件 word-pdf.ico 位于同一目录（例如 D:/word_to_PDF），使用以下指令打包为单个 .exe 文件：

bash

Copy
cd D:/word_to_PDF
pyinstaller -F -w --add-data "word-pdf.ico;." --icon=word-pdf.ico word_to_pdf_converter_office.py
-F：打包为单个 .exe 文件。
-w：以无控制台模式运行，关闭黑窗口。
--add-data "word-pdf.ico;."：添加图标文件（Windows 中分隔符为 ;）。
--icon=word-pdf.ico：内嵌图标到 .exe 文件。
3. 打包结果
打包完成后，.exe 文件会生成在 D:/word_to_PDF/dist 目录下，名为 word_to_pdf_converter_office.exe。
运行 .exe 文件时，日志会生成在 .exe 所在目录下的 converter.log。
使用方法
双击 word_to_pdf_converter_office.exe 启动程序。
点击“选择文件”按钮，选择需要转换的文件。
点击“开始转换”按钮，程序会自动转换文件。
转换完成后，PDF 文件会生成在原文件所在目录。
查看 converter.log 文件（位于 .exe 目录），了解转换详情。
常见问题
问题 1：程序提示“未检测到转换工具”
确保已安装 LibreOffice、Microsoft Office 或 WPS Office 至少一个。
确保已正确配置 PATH 环境变量。
问题 2：转换 .mhtml 文件失败
确保已安装 wkhtmltopdf，并配置 PATH。
检查 converter.log 中的错误信息，可能需要手动转换。
问题 3：转换 .pptx 文件失败
确保 Microsoft Office 或 WPS Office 已安装，且包含 PowerPoint 组件。
检查文件路径是否包含中文，建议重命名为纯英文路径，例如：
text

Copy
D:/word_to_PDF/test/1.pptx
获取帮助
如果您在使用过程中遇到问题，请联系开发者，提供以下信息：

converter.log 文件内容。
您的操作系统版本。
安装的 LibreOffice、wkhtmltopdf、Office/WPS 版本。
转换失败的文件（如果可以提供）。