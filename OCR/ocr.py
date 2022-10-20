import sys
from PIL import Image
import pyocr
import pyocr.builders
import xlsxwriter

tools = pyocr.get_available_tools()
if len(tools) == 0:
    print("No OCR tool found")
    sys.exit(1)

tool = tools[0]
print("Will use tool '%s'" % (tool.get_name()))


txt = tool.image_to_string(
    Image.open("D:\python\OCR\image.png"),
    lang="jpn",
    builder=pyocr.builders.TextBuilder(tesseract_layout=6)
)

line = 0
row = 0

workbook = xlsxwriter.Workbook("test.xlsx")
worksheet = workbook.add_worksheet("test")

worksheet.write(line, row, txt)
workbook.close()