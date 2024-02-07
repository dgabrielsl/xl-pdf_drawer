from win32com import client
import os

def xlpdf(file_location):

    os.system('taskkill /f /im excel.exe')

    app = client.DispatchEx('Excel.Application')
    app.Interactive = False
    app.Visible = False

    wb = app.Workbooks.open(file_location)

    # You must have indicate a location.
    output = r''

    wb.ActiveSheet.ExportAsFixedFormat(0, output)
    wb.Close()