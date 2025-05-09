Word to PDF Converter
概述
这是一个基于 Python 和 PyQt6 开发的批量文件转换工具，专门用于将 Microsoft Office 文件（包括 Word、PowerPoint、Excel 和 TXT 文件）转换为 PDF 格式。软件提供了一个简单的图形界面，允许用户选择多个文件并实时查看转换进度和日志。
功能

批量转换：支持将以下格式的文件批量转换为 PDF：
Word 文件：.doc、.docx
PowerPoint 文件：.ppt、.pptx
Excel 文件：.xls、.xlsx
文本文件：.txt（通过 Word 转换为 PDF）


实时 UI 更新：转换过程中，界面不会卡顿，实时显示进度和日志。
覆盖已存在文件：如果目标 PDF 文件已存在，程序会自动删除并覆盖。
错误重试：对于转换失败的文件，程序会自动重试（最多 3 次）。
依赖 Microsoft Office：使用 Microsoft Office 的 COM 接口进行转换，需要系统中安装 Microsoft Office（包括 Word、PowerPoint 和 Excel）。

运行要求
环境要求

操作系统：Windows（仅支持 Microsoft Office 的 COM 接口）
Microsoft Office：需要安装 Microsoft Office（包括 Word、PowerPoint 和 Excel），支持 Office 2010 及以上版本。
Python：Python 3.6 或以上版本。

依赖安装
程序需要以下 Python 库，使用国内镜像源（阿里云）加速下载：
1. 安装 PyQt6
pip install PyQt6 -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com

2. 安装 pywin32
pip install pywin32 -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com

3. 安装 pyinstaller（用于打包）
pip install pyinstaller -i https://mirrors.aliyun.com/pypi/simple/ --trusted-host mirrors.aliyun.com

注意

建议在虚拟环境中安装依赖，以避免冲突：python -m venv venv
venv\Scripts\activate  # Windows
# 或
source venv/bin/activate  # Linux/macOS



使用方法

运行程序：

确保已安装 Microsoft Office。
运行主脚本：python word_to_pdf_converter_office.py


界面会显示，选择需要转换的文件，点击“开始转换”。


功能说明：

点击“选择文件”按钮，选择需要转换的文件（支持多选）。
点击“开始转换”按钮，程序会异步执行转换，实时更新进度和日志。
如果目标 PDF 文件已存在，程序会自动覆盖。



打包为可执行文件
打包命令
将程序打包为单个 exe 文件，并内嵌 word-pdf.ico 图标，同时设置为程序图标：
pyinstaller --name WordToPDFConverter --onefile --windowed --icon=word-pdf.ico --add-data "word-pdf.ico;." word_to_pdf_converter_office.py

参数说明

--name WordToPDFConverter：设置可执行文件名称。
--onefile：打包为单个 exe 文件。
--windowed：以 GUI 模式运行（无控制台窗口）。
--icon=word-pdf.ico：设置程序图标。
--add-data "word-pdf.ico;."：内嵌图标文件（Windows 使用 ; 分隔符，Linux/macOS 使用 :）。
word_to_pdf_converter_office.py：主脚本文件名。

打包结果

打包完成后，可执行文件位于 dist/WordToPDFConverter.exe。
运行 WordToPDFConverter.exe，即可使用程序，无需安装 Python 环境（但仍需 Microsoft Office）。

注意事项

Microsoft Office 要求：
程序依赖 Microsoft Office 的 COM 接口，必须安装 Office（包括 Word、PowerPoint 和 Excel）。
如果 Office 未正确安装或激活，可能导致转换失败。


文件权限：
确保程序有权限访问输入文件和写入输出 PDF 文件的目录。
如果遇到权限问题，尝试以管理员身份运行程序。


图标文件：
确保 word-pdf.ico 文件存在于脚本同级目录，且为有效的 .ico 格式。


转换失败：
如果文件损坏或格式不兼容（例如 .doc 文件过旧），可能导致转换失败。
建议手动用 Word 打开文件，另存为 .docx 后再转换。



故障排查

转换失败：
检查日志中的错误信息，例如 Open.SaveAs 或 命令失败。
确保目标 PDF 文件未被占用（例如被 Adobe Acrobat 打开）。
手动用 Word 打开文件，确认文件是否损坏。


图标未显示：
确保 word-pdf.ico 文件有效。
检查打包命令中的 --add-data 参数是否正确。


Office 未检测到：
确保 Microsoft Office 已安装并激活。
修复 Office 安装：控制面板 -> 程序和功能 -> Microsoft Office -> 更改 -> 快速修复。



许可证
本项目仅供学习和个人使用，未指定具体许可证。使用 Microsoft Office 的 COM 接口需遵守 Microsoft 的许可协议。
