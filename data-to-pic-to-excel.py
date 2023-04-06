import xlsxwriter
import code128
from PIL import Image, ImageDraw, ImageFont

# Get barcode value
barcode_param = 'ATL-003111'

# Create barcode image
barcode_image = code128.image(barcode_param, height=120)

# Create empty image for barcode + text
top_bott_margin = 60
l_r_margin = 0
new_height = barcode_image.height + (2 * top_bott_margin)
new_width = barcode_image.width + (2 * l_r_margin)
new_image = Image.new('RGB', (new_width, new_height), (255, 255, 255))

# put barcode on new image
barcode_y = 50
new_image.paste(barcode_image, (0, barcode_y))

# object to draw text
draw = ImageDraw.Draw(new_image)

# Define custom text size and font
h1_size = 36

h1_font = ImageFont.truetype(
    "C:\\Windows\\Fonts\\Calibri\\Calibri\\calibri.ttf", h1_size)
# Define custom text
center_barcode_value = (barcode_image.width / 2) - len(barcode_param) * 8

# Draw text on picture
draw.text((center_barcode_value, (new_height - h1_size - 15)),
          barcode_param, fill=(0, 0, 0), font=h1_font)

workbook = xlsxwriter.Workbook(
    'path/to/excel/file.xlsx', {'constant_memory': True})
# save in file
new_image.save('barcode_image.png', 'PNG')

worksheet = workbook.add_worksheet()
worksheet.insert_image('B3', 'barcode_image.png') # specify cell here
workbook.close()


#-------------------------------------------------------------------------------------------------
import win32com.client
from pywintypes import com_error

# Path to original excel file
WB_PATH = 'C:\\Users\\gtsak\\Documents\\barcode-generator-and-excel-to-pdf\\New ATLs - dupe.xlsx'
# PDF path when saving
PATH_TO_PDF = 'C:\\Users\\gtsak\\Documents\\barcode-generator-and-excel-to-pdf\\New ATLs.pdf'
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
