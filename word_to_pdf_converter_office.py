import sys
import os
import time
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton,
    QFileDialog, QTextEdit, QLabel
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon
import win32com.client
import pythoncom

# 动态获取资源路径（支持 PyInstaller 的 --onefile 模式）
def resource_path(relative_path):
    """获取资源文件的绝对路径，支持 PyInstaller 的 --onefile 模式"""
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 创建的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

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

    def initialize_apps(self):
        try:
            self.word_app = win32com.client.Dispatch("Word.Application")
            self.word_app.Visible = False
        except Exception as e:
            self.update_log.emit(f"错误：无法初始化 Word！错误：{str(e)}\n")
            return False

        try:
            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        except Exception as e:
            self.update_log.emit(f"错误：无法初始化 PowerPoint，但将继续处理 Word 和 Excel 文件。错误：{str(e)}\n")
            self.ppt_app = None

        try:
            self.excel_app = win32com.client.Dispatch("Excel.Application")
            self.excel_app.Visible = False
        except Exception as e:
            self.update_log.emit(f"错误：无法初始化 Excel，但将继续处理 Word 和 PowerPoint 文件。错误：{str(e)}\n")
            self.excel_app = None

        return True

    def cleanup_apps(self):
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

    def run(self):
        pythoncom.CoInitialize()

        if not self.initialize_apps():
            self.conversion_finished.emit()
            pythoncom.CoUninitialize()
            return

        total_files = len(self.files)

        for i, file in enumerate(self.files, 1):
            retry_count = 0
            max_retries = 3
            success = False

            while retry_count < max_retries and not success:
                try:
                    if not os.path.exists(file):
                        raise FileNotFoundError("文件不存在或无法访问")

                    ext = os.path.splitext(file)[1].lower()
                    output_file = os.path.splitext(file)[0] + ".pdf"

                    # 检查目标 PDF 文件是否存在，如果存在则删除
                    if os.path.exists(output_file):
                        try:
                            os.remove(output_file)
                            self.update_log.emit(f"已删除现有文件：{output_file}\n")
                        except Exception as e:
                            raise RuntimeError(f"无法删除现有 PDF 文件 {output_file}，可能被占用。错误：{str(e)}")

                    if ext in [".doc", ".docx", ".txt"]:
                        if self.word_app is None:
                            raise RuntimeError("Word 应用程序不可用，无法转换此文件")
                        if ext == ".txt":
                            doc = self.word_app.Documents.Add()
                            doc.Content.Text = open(file, encoding="utf-8").read()
                        else:
                            doc = self.word_app.Documents.Open(file)
                        doc.SaveAs(output_file, FileFormat=17)
                        doc.Close()

                    elif ext in [".ppt", ".pptx"]:
                        if self.ppt_app is None:
                            raise RuntimeError("PowerPoint 应用程序不可用，无法转换此文件")
                        presentation = self.ppt_app.Presentations.Open(file, WithWindow=False)
                        presentation.SaveAs(output_file, FileFormat=32)
                        presentation.Close()

                    elif ext in [".xls", ".xlsx"]:
                        if self.excel_app is None:
                            raise RuntimeError("Excel 应用程序不可用，无法转换此文件")
                        workbook = self.excel_app.Workbooks.Open(file)
                        workbook.ExportAsFixedFormat(0, output_file)
                        workbook.Close()

                    else:
                        raise ValueError(f"不支持的文件格式：{ext}")

                    self.update_log.emit(f"[{i}/{total_files}] 转换成功：{file}\n")
                    success = True

                except Exception as e:
                    retry_count += 1
                    self.update_log.emit(f"[{i}/{total_files}] 转换失败（尝试 {retry_count}/{max_retries}）：{file}\n错误：{str(e)}\n")
                    if retry_count < max_retries:
                        if ext in [".doc", ".docx", ".txt"] and self.word_app:
                            try:
                                self.word_app.Quit()
                            except:
                                pass
                            self.word_app = win32com.client.Dispatch("Word.Application")
                            self.word_app.Visible = False
                        elif ext in [".ppt", ".pptx"] and self.ppt_app:
                            try:
                                self.ppt_app.Quit()
                            except:
                                pass
                            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                        elif ext in [".xls", ".xlsx"] and self.excel_app:
                            try:
                                self.excel_app.Quit()
                            except:
                                pass
                            self.excel_app = win32com.client.Dispatch("Excel.Application")
                            self.excel_app.Visible = False
                        time.sleep(1)
                    else:
                        self.update_log.emit(f"[{i}/{total_files}] 最终转换失败：{file}\n")

                # 更新进度
                progress = (i / total_files) * 100
                self.update_progress.emit(progress)

        self.cleanup_apps()
        self.conversion_finished.emit()
        pythoncom.CoUninitialize()

class WordToPDFConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.selected_files = []
        self.converter_thread = None

    def initUI(self):
        self.setWindowTitle("Word to PDF Converter")
        self.setGeometry(100, 100, 600, 400)

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

        self.status_label = QLabel("未选择文件")
        layout.addWidget(self.status_label)

        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)

    def check_office(self):
        try:
            pythoncom.CoInitialize()
            app = win32com.client.Dispatch("Word.Application")
            app.Visible = False
            app.Quit()
            return True
        except Exception as e:
            self.log_text.append(f"错误：未检测到 Microsoft Office！错误：{str(e)}\n")
            return False
        finally:
            pythoncom.CoUninitialize()

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "",
            "Supported Files (*.doc *.docx *.ppt *.pptx *.xls *.xlsx *.txt);;All Files (*.*)"
        )
        if files:
            self.selected_files = files
            self.convert_button.setEnabled(True)
            self.status_label.setText(f"已选择 {len(files)} 个文件")
            files_list = "\n".join(files)
            self.log_text.append(f"已选择文件：\n{files_list}\n")

    def start_conversion(self):
        if not self.selected_files:
            self.log_text.append("错误：未选择文件！\n")
            return

        if not self.check_office():
            self.log_text.append("请安装 Microsoft Office（包括 Word、PowerPoint 和 Excel）。\n")
            return

        self.convert_button.setEnabled(False)
        self.select_button.setEnabled(False)
        self.status_label.setText("转换中...")

        self.converter_thread = ConverterThread(self.selected_files)
        self.converter_thread.update_progress.connect(self.update_progress_label)
        self.converter_thread.update_log.connect(self.update_log_text)
        self.converter_thread.conversion_finished.connect(self.on_conversion_finished)
        self.converter_thread.start()

    def update_progress_label(self, progress):
        self.status_label.setText(f"转换进度：{progress:.1f}%")

    def update_log_text(self, message):
        self.log_text.append(message)

    def on_conversion_finished(self):
        self.status_label.setText("转换完成！")
        self.convert_button.setEnabled(True)
        self.select_button.setEnabled(True)
        self.log_text.append("所有文件转换完成！\n")

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