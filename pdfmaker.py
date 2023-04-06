import win32com.client
from pywintypes import com_error

# Path to original excel file
WB_PATH = 'path/to/excel/file'
# PDF path when saving
PATH_TO_PDF = 'path/to/pdf/file'
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
try:
    print('Start conversion to PDF')
    # Open
    wb = excel.Workbooks.Open(WB_PATH)
    # Specify the sheet you want to save by index. 1 is the first (leftmost) sheet.
    ws_index_list = [1]
    wb.WorkSheets(ws_index_list).Select()
    # Save
    wb.ActiveSheet.ExportAsFixedFormat(0, PATH_TO_PDF)
except com_error as e:
    print('failed.')
else:
    print('Succeeded.')
finally:
    wb.Close()
    excel.Quit()
