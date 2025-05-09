import win32com.client
import pythoncom

pythoncom.CoInitialize()
try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(r"D:/word_to_PDF/学生答辩记录表/1.黄紫晴.doc")
    doc.SaveAs(r"D:/word_to_PDF/学生答辩记录表/1.黄紫晴.pdf", FileFormat=17)
    doc.Close()
    word.Quit()
    print("转换成功")
except Exception as e:
    print(f"错误：{str(e)}")
finally:
    pythoncom.CoUninitialize()