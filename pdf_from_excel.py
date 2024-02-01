from win32com import client
import os

def xlpdf(file_location):
    app = client.DispatchEx('Excel.Application')
    app.Interactive = False
    app.Visible = False
    wb = app.Workbooks.open(file_location)
    output = r'C:\Users\dgabr\OneDrive\Documentos\Gabriel (cloud)\DeskPyLab\Lab - Projects for sale\XL-PDF drawer\Stylesheet - DeskPyLab.pdf'
    wb.ActiveSheet.ExportAsFixedFormat(0, output)
    wb.Close()

    try: os.system('taskkill /f /im excel.exe')
    except Exception as e: print(e)

xlpdf(r'C:\Users\dgabr\OneDrive\Documentos\Gabriel (cloud)\DeskPyLab\Lab - Projects for sale\XL-PDF drawer\Stylesheet.xlsx')