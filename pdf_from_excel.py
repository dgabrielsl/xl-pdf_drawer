from win32com import client
import os

def xlpdf(file_location):
    app = client.DispatchEx('Excel.Application')
    app.Interactive = False
    app.Visible = False
    wb = app.Workbooks.open(file_location)
    output = r'C:\Users\dgabr\OneDrive\Documentos\Gabriel (cloud)\DeskPyLab\Lab - Projects for sale\Log - DeskPyLab.pdf'
    wb.ActiveSheet.ExportAsFixedFormat(0, output)
    wb.Close()

xlpdf(r'C:\Users\dgabr\OneDrive\Documentos\Gabriel (cloud)\DeskPyLab\Lab - Projects for sale\Log.xlsx')