# Takes two positional arguments
# 1st - image location - mandatory
# 2nd - workbook name - optional

# TODO: Specify output folder of workbook


import sys
from datetime import date
from functools import reduce

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
    "Palam": (363, 326),
    "Lodhi Road": (361, 358),
    "Faridabad": (408, 382),
    "Ridge": (332, 365),
    "Ayanagar": (391, 340),
    "Najafgarh": (354, 305),
    "Ghaziabad": (339, 414),
    "Noida": (376, 397),
    "Gr. Noida": (383, 419),
    "Meerut": (254, 473),
    "Baghpat": (267, 359),
    "Bulandshahar": (408, 507),
    "Muzaffarnagar": (124, 473),
    "Gurugram": (395, 314),
    "Panipat": (146, 301)
}


# Maps color values in RGB format to rainfall range
rgb_to_rainfall = {
    (38, 38, 38): (93, 100),
    (41, 41, 41): (87,93),
    (42, 42, 42): (80,87),
    (43, 43, 43): (74,80),
    (44, 44, 44): (67,74),
    (40, 40, 40): (60,67),
    (45, 45, 45): (54,60),
    (23, 23, 23): (47,54),
    (16, 16, 16): (41,47),
    (5, 5, 5): (34,41),
    (4, 4, 4): (27,34),
    (3, 3, 3): (21,27),
    (2, 2, 2): (14,21),
    (1, 1, 1): (7.6,14),
    (6, 6, 6): (1.0,7.6)
}


# Find rainfall in places

result = []

for place, pos in places.items():
    # distance of observation points from position
    dist = 3

    # list of observation points around position
    points = []
    for i in range(-dist, dist + 1, dist):
        for j in range(-dist, dist + 1, dist):
            points.append((pos[0] + i, pos[1] + j))
    
    # list of RGB values at observation points
    rgb_values = []
    for point in points:
        # OpenCV stores color values in BGR format
        b, g, r = image[point[0], point[1]]
        rgb_values.append((r ,g, b))
    
    # print(rgb_values)

    # list of non-zero rain values (middle value of range)
    valid_rain_values = []
    for value in rgb_values:
        rain_low, rain_high = rgb_to_rainfall.get(value, (0, 0))
        rain_middle = (rain_low + rain_high) / 2
        if rain_middle != 0:
            valid_rain_values.append(rain_middle)

    # print(len(valid_rain_values))
    
    if len(valid_rain_values) == 0:
        result.append((place, "No rainfall"))
    else:
        rain_sum = reduce(lambda val1, val2: val1 + val2, valid_rain_values)
        rain_avg =  rain_sum / len(valid_rain_values)
        result.append((place, f'{"{0:.2f}".format(rain_avg - 2)} to {"{0:.2f}".format(rain_avg + 2)} mm rainfall'))


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
