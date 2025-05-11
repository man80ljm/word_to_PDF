import win32com.client
import pythoncom

pythoncom.CoInitialize()
try:
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    presentation = ppt.Presentations.Open(r"D:/word_to_PDF/测试/1.八仙过海盲盒玩偶设计.pptx", WithWindow=False)
    presentation.SaveAs(r"D:/word_to_PDF/测试/1.八仙过海盲盒玩偶设计.pdf", FileFormat=32)
    presentation.Close()
    ppt.Quit()
    print("转换成功")
except Exception as e:
    print(f"错误：{str(e)}")
finally:
    pythoncom.CoUninitialize()