# Takes two positional arguments
# 1st - image location - mandatory
# 2nd - workbook name - optional

# TODO: Specify output folder of workbook


import sys
from datetime import date
from functools import reduce

import xlsxwriter
from PIL import Image


if len(sys.argv) < 2:
    print("No input image. Please provide image location as argument")
    sys.exit()


# Read gif image using PIL/Pillow, and then convert into OpenCV format
image = Image.open(sys.argv[1], "r").convert("RGB")
image_pixels_rgb = list(image.getdata())
IMAGE_WIDTH = 880

# Stores location in (column, row) or (x, y) format
places = {
    "Palam": (326, 363),
    "Lodhi Road": (358, 361),
    "Ridge": (365, 332),
    "Ayanagar": (340, 391),
    "Najafgarh": (305, 354),
    "Faridabad": (382, 408),
    "Ghaziabad": (414, 339),
    "Noida": (397, 376),
    "Gr. Noida": (419, 383),
    "Gurugram": (314, 395),
    #
    "Muzaffarnagar": (473, 124),
    "Bijnor": (573, 150),
    "Meerut": (473, 254),
    "Baghpat": (359, 267),
    "Bulandshahar": (507, 408),
    "Sambhal": (675, 360),
    "Aligarh": (560, 545),
    "Hathras": (556, 623),
    "Mathura": (467, 653),
    #
    "Panipat": (301, 146),
    "Karnal": (306, 67),
    "Kaithal": (169, 38),
    "Jind": (149, 165),
    "Hissar": (10, 209),
    "Rohtak": (213, 279),
    "Bhiwani": (106, 303),
    "Jhajjar": (227, 357),
    "Rewari": (216, 467),
    "Narnaul": (96, 499),
    "Nuh": (309, 490),
    "Palwal": (385, 480)
}


# Maps color values in RGB format to rainfall range
rgb_to_rainfall = {
    (200, 0, 0): (93, 100),
    (255, 63, 0): (87, 93),
    (255, 115, 0): (80, 87),
    (255, 189, 0): (74, 80),
    (255, 230, 0): (67, 74),
    (252, 252, 112): (60, 67),
    (255, 255, 255): (54, 60),
    (135, 241, 255): (47, 54),
    (83, 209, 255): (41, 47),
    (26, 163, 255): (34, 41),
    (0, 121, 255): (27, 34),
    (0, 71, 255): (21, 27),
    (0, 58, 200): (14, 21),
    (0, 25, 176): (7.6, 14),
    (58, 0, 160): (1.0, 7.6)
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
        index = (point[1] - 1) * IMAGE_WIDTH + point[0]
        rgb_values.append(image_pixels_rgb[index])

    # list of non-zero rain values (middle value of range)
    valid_rain_values = []
    for value in rgb_values:
        rain_low, rain_high = rgb_to_rainfall.get(value, (0, 0))
        rain_middle = (rain_low + rain_high) / 2
        if rain_middle != 0:
            valid_rain_values.append(rain_middle)

    if len(valid_rain_values) == 0:
        result.append((place, "No rainfall"))
    else:
        rain_sum = reduce(lambda val1, val2: val1 + val2, valid_rain_values)
        rain_avg = rain_sum / len(valid_rain_values)

        # Find rainfall range in which rain_avg lies
        rainfall_ranges = list(rgb_to_rainfall.values())
        rainfall_ranges.reverse()
        low, high = -1, -1

        for rainfall_range in rainfall_ranges:
            if rain_avg < rainfall_range[1]:
                low = rainfall_range[0]
                high = rainfall_range[1]
                break

        result.append(
            (place, f'{low} to {high} mm rainfall'))


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
