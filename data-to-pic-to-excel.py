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
    "C:\Windows\Fonts\Calibri\Calibri\calibri.ttf", h1_size)
# Define custom text
center_barcode_value = (barcode_image.width / 2) - len(barcode_param) * 8

# Draw text on picture
draw.text((center_barcode_value, (new_height - h1_size - 15)),
          barcode_param, fill=(0, 0, 0), font=h1_font)

workbook = xlsxwriter.Workbook(
    'path/to/excel/file.xlsx', {'constant_memory': True})
worksheet = workbook.add_worksheet()
worksheet.insert_image('B3', 'barcode_image.png') # specify cell here
workbook.close()
