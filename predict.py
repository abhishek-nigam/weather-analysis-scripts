import sys
from datetime import date

import numpy
import xlsxwriter
from PIL import Image
from cv2 import cv2


if len(sys.argv) < 2:
    print("No input image. Please provide image location as argument")
    sys.exit()


# Read gif image using PIL/Pillow, and then convert into OpenCV format
pil_image = Image.open(sys.argv[1])
image = cv2.cvtColor(numpy.array(pil_image), cv2.COLOR_RGB2BGR)


# Stores location in (row, column) format
# NOTE: If using conventional map (x,y) coordinates,
# NOTE: write them as (y,x) here
places = {
    "Place 1": (432, 166),
    "Place 2": (511, 188),
}


# Maps color values in RGB format to rainfall
rgb_to_rainfall = {
    (38, 38, 38): '93 - 100 mm',
    (41, 41, 41): '87 - 93 mm',
    (42, 42, 42): '80 - 87 mm',
    (43, 43, 43): '74 - 80 mm',
    (44, 44, 44): '67 - 74 mm',
    (40, 40, 40): '60 - 67 mm',
    (45, 45, 45): '54 - 60 mm',
    (23, 23, 23): '47 - 54 mm',
    (16, 16, 16): '41 - 47 mm',
    (5, 5, 5): '34 - 41 mm',
    (4, 4, 4): '27 - 34 mm',
    (3, 3, 3): '21 - 27 mm',
    (2, 2, 2): '14 - 21 mm',
    (1, 1, 1): '7.6 - 14 mm',
    (6, 6, 6): '1.0 - 7.6 mm'
}


# Find rainfall in places

result = []

for place, pos in places.items():
    # Open CV stores color values in BGR format
    b, g, r = image[pos[0], pos[1]]
    # Rearranging color values, and obtaining rainfall
    rainfall = rgb_to_rainfall[(r, g, b)]
    result.append((place, rainfall))


# Write result in Workbook

workbook_name = None

if len(sys.argv) >= 3:
    # Accepts workbook name as second argument
    workbook_name = sys.argv[2]
else:
    # Defaults to current date e.g. 2019-06-19
    workbook_name = str(date.today())


workbook = xlsxwriter.Workbook(workbook_name + ".xlsx")
worksheet = workbook.add_worksheet()
# Widens first and second column
worksheet.set_column('A:A', 20)
worksheet.set_column('B:B', 30)

# Add column header
format_bold = workbook.add_format({'bold': True})
worksheet.write(0, 0, "City", format_bold)
worksheet.write(0, 1, "Rainfall", format_bold)

for count, data in enumerate(result, 1):
    # Writes place in first, and rainfall in second column
    worksheet.write(count, 0, data[0])
    worksheet.write(count, 1, data[1])

workbook.close()
