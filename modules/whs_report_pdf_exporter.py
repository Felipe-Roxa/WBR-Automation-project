import os
import win32com.client
import time
import psutil

def pdf_exporter(xlsx_simple_path):
    """This function read a .xlsx file and export it as .pdf file"""

    xlsx_abs_path = os.path.abspath(xlsx_simple_path)
    lenght_xlsx_simple_path = len(xlsx_simple_path)
    pdf_simple_path = xlsx_simple_path[0:lenght_xlsx_simple_path - 4] + 'pdf'

    pdf_simple_path = 'Output/WHS Weekly Metrics Report.pdf'
    pdf_abs_path = os.path.abspath(pdf_simple_path)

    xlapp = win32com.client.DispatchEx("Excel.Application")
    wb = xlapp.Workbooks.Open(xlsx_abs_path)
    wb.RefreshAll()
    time.sleep(5)
    xlapp.CalculateUntilAsyncQueriesDone()
    xlapp.DisplayAlerts = False
    wb.Save()
    wb.Close(SaveChanges=True)
    xlapp.Quit()

    wb = xlapp.Workbooks.Open(xlsx_abs_path)
    wb.RefreshAll()
    ws_index_list = [1]
    xlapp.Visible = False
    wb.WorkSheets(ws_index_list).Select()
    wb.ActiveSheet.ExportAsFixedFormat(0, pdf_abs_path)
    xlapp.DisplayAlerts = False
    wb.Save()
    wb.Close(SaveChanges=True)
    xlapp.Quit()

    for proc in psutil.process_iter():
        if proc.name() == "excel.exe":
            proc.kill()
        else:
            continue