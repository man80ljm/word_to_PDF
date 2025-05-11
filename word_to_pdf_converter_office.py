import sys
import os
import time
import subprocess
import win32com.client
import pythoncom
import win32api
import logging
import pdfkit
import email
import base64
import tempfile
import re
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QProgressBar, QLabel
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QPropertyAnimation
from PyQt6.QtGui import QIcon

# 配置日志，保存到程序同级目录下的 converter.log
if hasattr(sys, 'frozen'):
    script_dir = os.path.dirname(os.path.abspath(sys.executable))
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))
log_file = os.path.join(script_dir, "converter.log")
logging.basicConfig(
    filename=log_file,
    filemode="w",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger()

# 动态获取资源路径（支持 PyInstaller 的 --onefile 模式）
def resource_path(relative_path):
    """获取资源文件的绝对路径，支持 PyInstaller 的 --onefile 模式"""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def extract_mhtml_to_html(mhtml_file):
    """解析 .mhtml 文件，提取 HTML 内容和嵌入资源，保存为临时文件"""
    try:
        with open(mhtml_file, 'rb') as f:
            msg = email.message_from_binary_file(f)
        
        html_content = None
        resources = {}
        temp_dir = tempfile.gettempdir()
        
        # 提取 HTML 和嵌入资源
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/html":
                charset = part.get_content_charset() or "utf-8"
                html_content = part.get_payload(decode=True).decode(charset, errors="replace")
            elif content_type.startswith("image/"):
                content_id = part.get("Content-ID", "").strip("<>")
                if content_id:
                    # 提取图片数据
                    img_data = part.get_payload(decode=True)
                    img_ext = content_type.split("/")[-1]  # 例如，image/png → png
                    img_filename = f"resource_{content_id}.{img_ext}"
                    img_path = os.path.join(temp_dir, img_filename)
                    with open(img_path, "wb") as f:
                        f.write(img_data)
                    resources[content_id] = img_filename
        
        if not html_content:
            return None, "未找到 HTML 内容"
        
        # 替换 HTML 中的资源引用（cid:xxx → 本地文件路径）
        for content_id, filename in resources.items():
            cid_url = f"cid:{content_id}"
            local_url = filename
            html_content = html_content.replace(cid_url, local_url)
        
        # 创建临时文件保存 HTML 内容
        temp_html_file = os.path.join(temp_dir, f"mhtml_temp_{os.path.basename(mhtml_file)}.html")
        with open(temp_html_file, "w", encoding="utf-8") as f:
            f.write(html_content)
        
        return temp_html_file, None
    except Exception as e:
        return None, f"解析 .mhtml 文件失败：{str(e)}"

def convert_to_pdf_wkhtmltopdf(input_file, output_dir):
    """使用 wkhtmltopdf 转换网页文件到 PDF"""
    output_file = os.path.join(output_dir, os.path.splitext(os.path.basename(input_file))[0] + ".pdf")
    temp_html_file = None
    try:
        if not os.path.exists(input_file):
            raise FileNotFoundError("文件不存在或无法访问")
        
        # 处理中文路径
        short_input = win32api.GetShortPathName(input_file)
        short_output = win32api.GetShortPathName(os.path.dirname(input_file)) + "\\" + os.path.splitext(os.path.basename(input_file))[0] + ".pdf"
        output_file = os.path.join(output_dir, os.path.splitext(os.path.basename(input_file))[0] + ".pdf")
        
        # 检查目标 PDF 文件是否存在，如果存在则删除
        if os.path.exists(output_file):
            try:
                os.remove(output_file)
                logger.info(f"已删除现有文件：{output_file}")
            except Exception as e:
                raise RuntimeError(f"无法删除现有 PDF 文件 {output_file}，可能被占用。错误：{str(e)}")
        
        # 如果是 .mhtml 文件，先提取 HTML 和资源
        ext = os.path.splitext(input_file)[1].lower()
        if ext == ".mhtml":
            temp_html_file, error = extract_mhtml_to_html(input_file)
            if temp_html_file is None:
                raise ValueError(error)
            short_input = win32api.GetShortPathName(temp_html_file)
        
        # 指定 wkhtmltopdf 路径（根据实际安装路径修改）
        wkhtmltopdf_path = "C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
        if not os.path.exists(wkhtmltopdf_path):
            raise FileNotFoundError(f"未找到 wkhtmltopdf 可执行文件：{wkhtmltopdf_path}")
        
        # 配置 pdfkit，使用指定的 wkhtmltopdf 路径和选项
        config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
        options = {
            "--enable-local-file-access": "",
            "--encoding": "UTF-8",
            "--no-stop-slow-scripts": "",
            "--javascript-delay": "2000"
        }
        pdfkit.from_file(short_input, short_output, configuration=config, options=options)
        
        if not os.path.exists(output_file):
            raise RuntimeError("wkhtmltopdf 转换失败，未生成 PDF 文件")
        
        return output_file, None
    except Exception as e:
        return None, f"wkhtmltopdf 错误：{str(e)}"
    finally:
        # 清理临时文件
        if temp_html_file and os.path.exists(temp_html_file):
            try:
                os.remove(temp_html_file)
            except:
                pass
        # 清理临时资源文件
        temp_dir = tempfile.gettempdir()
        for filename in os.listdir(temp_dir):
            if filename.startswith("resource_") and (filename.endswith(".png") or filename.endswith(".jpg") or filename.endswith(".jpeg") or filename.endswith(".gif")):
                try:
                    os.remove(os.path.join(temp_dir, filename))
                except:
                    pass

def convert_to_pdf_libreoffice(input_file, output_dir):
    """使用 LibreOffice 转换文件到 PDF"""
    output_file = os.path.join(output_dir, os.path.splitext(os.path.basename(input_file))[0] + ".pdf")
    try:
        if not os.path.exists(input_file):
            raise FileNotFoundError("文件不存在或无法访问")
        
        # 检查目标 PDF 文件是否存在，如果存在则删除
        if os.path.exists(output_file):
            try:
                os.remove(output_file)
                logger.info(f"已删除现有文件：{output_file}")
            except Exception as e:
                raise RuntimeError(f"无法删除现有 PDF 文件 {output_file}，可能被占用。错误：{str(e)}")
        
        # 指定 LibreOffice soffice 路径（根据实际安装路径修改）
        soffice_path = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
        if not os.path.exists(soffice_path):
            raise FileNotFoundError(f"未找到 LibreOffice 可执行文件：{soffice_path}")
        
        # 使用 LibreOffice 命令行转换，隐藏 DOS 窗口，添加 --norestore 和 --safe-mode
        result = subprocess.run([
            soffice_path, "--headless", "--convert-to", "pdf",
            "--norestore", "--safe-mode",
            input_file, "--outdir", output_dir
        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)
        
        # 记录 LibreOffice 输出
        if result.stdout:
            logger.debug(f"LibreOffice stdout: {result.stdout}")
        if result.stderr:
            logger.warning(f"LibreOffice stderr: {result.stderr}")
        
        if not os.path.exists(output_file):
            raise RuntimeError(f"LibreOffice 转换失败，未生成 PDF 文件。详情：{result.stderr}")
        
        return output_file, None
    except Exception as e:
        return None, f"LibreOffice 错误：{str(e)}"

def check_libreoffice():
    """检查 LibreOffice 是否可用"""
    try:
        # 指定 LibreOffice soffice 路径
        soffice_path = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"
        result = subprocess.run([soffice_path, "--version"], check=True,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)
        return True, result.stdout.strip()
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        return False, f"未找到 LibreOffice 或无法运行：{str(e)}"

def check_wkhtmltopdf():
    """检查 wkhtmltopdf 是否可用"""
    try:
        # 指定 wkhtmltopdf 路径
        wkhtmltopdf_path = "C:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe"
        result = subprocess.run([wkhtmltopdf_path, "--version"], check=True,
                                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True,
                                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0)
        return True, result.stdout.strip()
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        return False, f"未找到 wkhtmltopdf 或无法运行：{str(e)}"

class ConverterThread(QThread):
    update_progress = pyqtSignal(float)
    update_log = pyqtSignal(str)
    conversion_finished = pyqtSignal()

    def __init__(self, files):
        super().__init__()
        self.files = files
        self.word_app = None
        self.ppt_app = None
        self.excel_app = None
        self.wps_ppt_app = None

    def initialize_office_apps(self):
        """初始化 Microsoft Office 应用程序"""
        try:
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
        except Exception as e:
            logger.error(f"无法初始化 Office Word：{str(e)}")
            return False

        try:
            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as e:
            logger.warning(f"无法初始化 Office PowerPoint，将尝试 WPS：{str(e)}")
            self.ppt_app = None

        try:
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = False
        except Exception as e:
            logger.warning(f"无法初始化 Office Excel，将继续处理其他文件：{str(e)}")
            self.excel_app = None

        return self.word_app is not None or self.ppt_app is not None or self.excel_app is not None

    def initialize_wps_apps(self):
        """初始化 WPS 应用程序"""
        try:
            self.wps_ppt_app = win32com.client.Dispatch("KWPP.Application")
        except Exception as e:
            logger.warning(f"无法初始化 WPS 演示：{str(e)}")
            self.wps_ppt_app = None

        return self.wps_ppt_app is not None

    def cleanup_office_apps(self):
        """清理 Microsoft Office 应用程序"""
        try:
            if self.word_app:
                self.word_app.Quit()
        except:
            pass
        try:
            if self.ppt_app:
                self.ppt_app.Quit()
        except:
            pass
        try:
            if self.excel_app:
                self.excel_app.Quit()
        except:
            pass
        self.word_app = None
        self.ppt_app = None
        self.excel_app = None

    def cleanup_wps_apps(self):
        """清理 WPS 应用程序"""
        try:
            if self.wps_ppt_app:
                self.wps_ppt_app.Quit()
        except:
            pass
        self.wps_ppt_app = None

    def convert_with_office(self, file):
        """使用 Microsoft Office COM 接口转换文件"""
        try:
            if not os.path.exists(file):
                raise FileNotFoundError("文件不存在或无法访问")

            ext = os.path.splitext(file)[1].lower()
            output_file = os.path.join(os.path.dirname(file), os.path.splitext(os.path.basename(file))[0] + ".pdf")

            # 处理中文路径
            short_file = win32api.GetShortPathName(file)
            short_output = win32api.GetShortPathName(os.path.dirname(file)) + "\\" + os.path.splitext(os.path.basename(file))[0] + ".pdf"

            # 检查目标 PDF 文件是否存在，如果存在则删除
            if os.path.exists(output_file):
                try:
                    os.remove(output_file)
                    logger.info(f"已删除现有文件：{output_file}")
                except Exception as e:
                    raise RuntimeError(f"无法删除现有 PDF 文件 {output_file}，可能被占用。错误：{str(e)}")

            if ext in [".doc", ".docx", ".txt"]:
                if self.word_app is None:
                    raise RuntimeError("Office Word 应用程序不可用")
                if ext == ".txt":
                    doc = self.word_app.Documents.Add()
                    doc.Content.Text = open(file, encoding="utf-8").read()
                else:
                    doc = self.word_app.Documents.Open(short_file)
                doc.SaveAs(short_output, FileFormat=17)
                doc.Close()
                return output_file, None

            elif ext in [".ppt", ".pptx"]:
                if self.ppt_app is None:
                    raise RuntimeError("Office PowerPoint 应用程序不可用")
                presentation = self.ppt_app.Presentations.Open(short_file, WithWindow=False)
                presentation.SaveAs(short_output, FileFormat=32)
                presentation.Close()
                return output_file, None

            elif ext in [".xls", ".xlsx"]:
                if self.excel_app is None:
                    raise RuntimeError("Office Excel 应用程序不可用")
                workbook = self.excel_app.Workbooks.Open(short_file)
                workbook.ExportAsFixedFormat(0, short_output)
                workbook.Close()
                return output_file, None

            else:
                raise ValueError(f"不支持的文件格式：{ext}")

        except Exception as e:
            return None, f"Office COM 错误：{str(e)}"

    def convert_with_wps(self, file):
        """使用 WPS COM 接口转换文件（仅 PPT）"""
        try:
            if not os.path.exists(file):
                raise FileNotFoundError("文件不存在或无法访问")

            ext = os.path.splitext(file)[1].lower()
            if ext not in [".ppt", ".pptx"]:
                raise ValueError(f"WPS 仅支持 PPT/PPTX 文件，当前格式：{ext}")

            output_file = os.path.join(os.path.dirname(file), os.path.splitext(os.path.basename(file))[0] + ".pdf")

            # 处理中文路径
            short_file = win32api.GetShortPathName(file)
            short_output = win32api.GetShortPathName(os.path.dirname(file)) + "\\" + os.path.splitext(os.path.basename(file))[0] + ".pdf"

            # 检查目标 PDF 文件是否存在，如果存在则删除
            if os.path.exists(output_file):
                try:
                    os.remove(output_file)
                    logger.info(f"已删除现有文件：{output_file}")
                except Exception as e:
                    raise RuntimeError(f"无法删除现有 PDF 文件 {output_file}，可能被占用。错误：{str(e)}")

            if self.wps_ppt_app is None:
                raise RuntimeError("WPS 演示应用程序不可用")

            presentation = self.wps_ppt_app.Presentations.Open(short_file, WithWindow=False)
            presentation.SaveAs(short_output, FileFormat=32)
            presentation.Close()
            return output_file, None

        except Exception as e:
            return None, f"WPS COM 错误：{str(e)}"

    def run(self):
        pythoncom.CoInitialize()
        libreoffice_available, libreoffice_status = check_libreoffice()
        wkhtmltopdf_available, wkhtmltopdf_status = check_wkhtmltopdf()
        if libreoffice_available:
            logger.info(f"检测到 LibreOffice：{libreoffice_status}")
        else:
            logger.error(f"LibreOffice 不可用：{libreoffice_status}")
            self.update_log.emit("错误：LibreOffice 不可用，请检查安装或 PATH 配置！")
        if wkhtmltopdf_available:
            logger.info(f"检测到 wkhtmltopdf：{wkhtmltopdf_status}")
        else:
            logger.warning(f"wkhtmltopdf 不可用：{wkhtmltopdf_status}")
            self.update_log.emit("警告：wkhtmltopdf 不可用，网页文件转换可能失败！请安装 wkhtmltopdf 或手动转换网页文件。")

        office_initialized = False
        wps_initialized = False
        total_files = len(self.files)

        for i, file in enumerate(self.files, 1):
            ext = os.path.splitext(file)[1].lower()
            output_dir = os.path.dirname(file)
            success = False
            error_msg = ""

            logger.info(f"开始转换文件 [{i}/{total_files}]：{file}")

            # 模拟细粒度进度更新
            for step in range(1, 11):
                sub_progress = (i - 1 + step / 10) / total_files * 100
                self.update_progress.emit(sub_progress)
                time.sleep(0.2)

            # 1. 对于 .html 和 .mhtml 文件，优先尝试 wkhtmltopdf
            if ext in [".html", ".mhtml"]:
                if wkhtmltopdf_available:
                    retry_count = 0
                    max_retries = 3
                    while retry_count < max_retries and not success:
                        try:
                            output_file, error = convert_to_pdf_wkhtmltopdf(file, output_dir)
                            if output_file:
                                logger.info(f"[{i}/{total_files}] wkhtmltopdf 转换成功：{file}")
                                success = True
                            else:
                                raise Exception(error)
                        except Exception as e:
                            retry_count += 1
                            error_msg = str(e)
                            logger.error(f"[{i}/{total_files}] wkhtmltopdf 转换失败（尝试 {retry_count}/{max_retries}）：{file}，错误：{error_msg}")
                            if retry_count < max_retries:
                                time.sleep(1)

            # 2. 如果 wkhtmltopdf 失败或不可用，尝试 LibreOffice（对于 .html 和 .mhtml）
            if not success and ext in [".html", ".mhtml"]:
                if libreoffice_available:
                    retry_count = 0
                    max_retries = 3
                    while retry_count < max_retries and not success:
                        try:
                            output_file, error = convert_to_pdf_libreoffice(file, output_dir)
                            if output_file:
                                logger.info(f"[{i}/{total_files}] LibreOffice 转换成功：{file}")
                                success = True
                            else:
                                raise Exception(error)
                        except Exception as e:
                            retry_count += 1
                            error_msg = str(e)
                            logger.error(f"[{i}/{total_files}] LibreOffice 转换失败（尝试 {retry_count}/{max_retries}）：{file}，错误：{error_msg}")
                            if retry_count < max_retries:
                                time.sleep(1)

            # 3. 对于非网页文件，尝试 LibreOffice
            if not success and ext not in [".html", ".mhtml"]:
                if libreoffice_available:
                    retry_count = 0
                    max_retries = 3
                    while retry_count < max_retries and not success:
                        try:
                            output_file, error = convert_to_pdf_libreoffice(file, output_dir)
                            if output_file:
                                logger.info(f"[{i}/{total_files}] LibreOffice 转换成功：{file}")
                                success = True
                            else:
                                raise Exception(error)
                        except Exception as e:
                            retry_count += 1
                            error_msg = str(e)
                            logger.error(f"[{i}/{total_files}] LibreOffice 转换失败（尝试 {retry_count}/{max_retries}）：{file}，错误：{error_msg}")
                            if retry_count < max_retries:
                                time.sleep(1)

            # 4. 如果 LibreOffice 失败，尝试 Microsoft Office（仅 Windows）
            if not success and ext not in [".html", ".mhtml"] and sys.platform == "win32":
                if not office_initialized:
                    office_initialized = self.initialize_office_apps()
                
                if office_initialized:
                    retry_count = 0
                    max_retries = 3
                    while retry_count < max_retries and not success:
                        try:
                            output_file, error = self.convert_with_office(file)
                            if output_file:
                                logger.info(f"[{i}/{total_files}] Office COM 转换成功：{file}")
                                success = True
                            else:
                                raise Exception(error)
                        except Exception as e:
                            retry_count += 1
                            error_msg = str(e)
                            logger.error(f"[{i}/{total_files}] Office COM 转换失败（尝试 {retry_count}/{max_retries}）：{file}，错误：{error_msg}")
                            if retry_count < max_retries:
                                self.cleanup_office_apps()
                                office_initialized = self.initialize_office_apps()
                                time.sleep(1)

            # 5. 如果 Office 失败且是 PPT/PPTX，尝试 WPS（仅 Windows）
            if not success and ext in [".ppt", ".pptx"] and sys.platform == "win32":
                if not wps_initialized:
                    wps_initialized = self.initialize_wps_apps()
                
                if wps_initialized:
                    retry_count = 0
                    max_retries = 3
                    while retry_count < max_retries and not success:
                        try:
                            output_file, error = self.convert_with_wps(file)
                            if output_file:
                                logger.info(f"[{i}/{total_files}] WPS COM 转换成功：{file}")
                                success = True
                            else:
                                raise Exception(error)
                        except Exception as e:
                            retry_count += 1
                            error_msg = str(e)
                            logger.error(f"[{i}/{total_files}] WPS COM 转换失败（尝试 {retry_count}/{max_retries}）：{file}，错误：{error_msg}")
                            if retry_count < max_retries:
                                self.cleanup_wps_apps()
                                wps_initialized = self.initialize_wps_apps()
                                time.sleep(1)

            if not success:
                logger.error(f"[{i}/{total_files}] 最终转换失败：{file}，错误：{error_msg}")
                self.update_log.emit(f"转换失败：{os.path.basename(file)}")

        if office_initialized:
            self.cleanup_office_apps()
        if wps_initialized:
            self.cleanup_wps_apps()
        self.conversion_finished.emit()
        pythoncom.CoUninitialize()

class WordToPDFConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []
        self.converter_thread = None
        self.current_progress = 0
        self.progress_animation = None

    def initUI(self):
        self.setWindowTitle("Word to PDF Converter")
        self.setGeometry(600, 400, 600, 150)

        # 动态加载图标
        icon_path = resource_path("word-pdf.ico")
        self.setWindowIcon(QIcon(icon_path))

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        self.select_button = QPushButton("选择文件")
        self.select_button.clicked.connect(self.select_files)
        layout.addWidget(self.select_button)

        self.convert_button = QPushButton("开始转换")
        self.convert_button.clicked.connect(self.start_conversion)
        self.convert_button.setEnabled(False)
        layout.addWidget(self.convert_button)

        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 2px solid #4CAF50;
                border-radius: 10px;
                text-align: center;
                height: 20px;
                background-color: #f0f0f0;
            }
            QProgressBar::chunk {
                background-color: #1B5E20; /* 墨绿色 */
                border-radius: 8px;
            }
        """)
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("未选择文件")
        layout.addWidget(self.status_label)

    def check_requirements(self):
        """检查 LibreOffice、Microsoft Office 和 WPS 是否可用"""
        libreoffice_ok, libreoffice_status = check_libreoffice()
        wkhtmltopdf_ok, wkhtmltopdf_status = check_wkhtmltopdf()
        office_ok = False
        wps_ok = False

        if sys.platform == "win32":
            try:
                pythoncom.CoInitialize()
                app = win32com.client.Dispatch("Word.Application")
                app.Visible = False
                app.Quit()
                office_ok = True
            except Exception as e:
                logger.warning(f"未检测到 Microsoft Office Word：{str(e)}")
            finally:
                pythoncom.CoUninitialize()

            try:
                pythoncom.CoInitialize()
                app = win32com.client.Dispatch("KWPP.Application")
                app.Quit()
                wps_ok = True
            except Exception as e:
                logger.warning(f"未检测到 WPS 演示：{str(e)}")
            finally:
                pythoncom.CoUninitialize()
        
        if not libreoffice_ok and not office_ok and not wps_ok and not wkhtmltopdf_ok:
            logger.error("未检测到 LibreOffice、Microsoft Office、WPS 或 wkhtmltopdf")
            self.status_label.setText("错误：未检测到转换工具！")
            return False
        
        logger.info(f"环境检查：LibreOffice={'可用' if libreoffice_ok else '不可用'}，wkhtmltopdf={'可用' if wkhtmltopdf_ok else '不可用'}，Office={'可用' if office_ok else '不可用'}，WPS={'可用' if wps_ok else '不可用'}")
        return True

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "",
            "Supported Files (*.doc *.docx *.ppt *.pptx *.xls *.xlsx *.txt *.html *.mhtml);;All Files (*.*)"
        )
        if files:
            self.selected_files = files
            self.convert_button.setEnabled(True)
            self.status_label.setText(f"已选择 {len(files)} 个文件")
            files_list = "\n".join(files)
            logger.info(f"已选择文件：\n{files_list}")

    def start_conversion(self):
        if not self.selected_files:
            self.status_label.setText("错误：未选择文件！")
            logger.error("未选择文件")
            return

        if not self.check_requirements():
            return

        self.convert_button.setEnabled(False)
        self.select_button.setEnabled(False)
        self.status_label.setText("转换中...")
        self.current_progress = 0
        self.progress_bar.setValue(0)

        self.converter_thread = ConverterThread(self.selected_files)
        self.converter_thread.update_progress.connect(self.update_progress_bar)
        self.converter_thread.update_log.connect(self.update_status_label)
        self.converter_thread.conversion_finished.connect(self.on_conversion_finished)
        self.converter_thread.start()

    def update_progress_bar(self, progress):
        # 使用 QPropertyAnimation 实现平滑过渡
        if self.progress_animation:
            self.progress_animation.stop()

        target_progress = int(progress)
        logger.debug(f"进度条动画：从 {self.current_progress} 到 {target_progress}")
        self.progress_animation = QPropertyAnimation(self.progress_bar, b"value")
        self.progress_animation.setDuration(1000)  # 动画持续时间 1000ms
        self.progress_animation.setStartValue(self.current_progress)
        self.progress_animation.setEndValue(target_progress)
        self.progress_animation.start()

        self.current_progress = target_progress
        self.status_label.setText(f"转换进度：{progress:.1f}%")

    def update_status_label(self, message):
        self.status_label.setText(message)

    def on_conversion_finished(self):
        # 确保进度条平滑过渡到 100%
        self.update_progress_bar(100)
        self.status_label.setText("转换完成！")
        self.convert_button.setEnabled(True)
        self.select_button.setEnabled(True)
        logger.info("所有文件转换完成")

def main():
    app = QApplication(sys.argv)
    # 动态加载图标
    icon_path = resource_path("word-pdf.ico")
    app.setWindowIcon(QIcon(icon_path))
    window = WordToPDFConverter()
    window.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()