import configparser
import json
import math
import os
import time

import numpy as np
from colormath.color_conversions import convert_color
from colormath.color_diff import delta_e_cie2000
from colormath.color_objects import LabColor, sRGBColor
from PIL import Image
import xlsxwriter

img_name = 'ultraball.png'
dmc_map = 'all'  # all or owned
manual_override = {
    (48, 48, 48, 255): 3799
}

input_folder = 'in/'
output_folder = 'out/'
img_path = input_folder + img_name

start_time = time.time()
os.chdir(os.path.dirname(__file__))
# print(os.getcwd())

config = configparser.ConfigParser()
config.read('config.ini')
all_dmc = dict(list(config['ALL'].items()))
owned_dmc = dict(list(config['OWNED'].items()))

for key in all_dmc.keys():
    all_dmc[key] = eval(all_dmc[key])

for key in owned_dmc.keys():
    owned_dmc[key] = eval(owned_dmc[key])

if dmc_map == 'all':
    dmc_map = all_dmc
else:
    dmc_map = owned_dmc

chars = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j',
                        'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
                        'u', 'v', 'w', 'x', 'y', 'z', 'A', 'B', 'C', 'D',
                        'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N',
                        'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X',
                        'Y', 'Z',
                        '1', '2', '3', '4', '5', '6', '7', '8', '9', '0']
color_counter = 0

color_to_letter_map = {}

def get_closest_color(pixel_rgb, dmc_dict, method='Richard'):

    min_diff = 9999999
    min_dmc = ''
    # Search for overrides
    if tuple(pixel_rgb) in manual_override:
        return str(manual_override[tuple(pixel_rgb)])

    if method == 'Richard':
        for dmc in dmc_dict:
            dmc_rgb = dmc_dict[dmc]['rgb']
            r_line = (pixel_rgb[0] + dmc_rgb[0]) / 2

            r_w_diff = (2 + (r_line / 256.0)) * \
                ((dmc_rgb[0] - pixel_rgb[0]) ** 2)
            g_w_diff = 4 * ((dmc_rgb[1] - pixel_rgb[1]) ** 2)
            b_w_diff = (2 + ((255 - r_line) / 256)) * \
                ((dmc_rgb[2] - pixel_rgb[2]) ** 2)
            color_diff = math.sqrt(r_w_diff + g_w_diff + b_w_diff)

            if color_diff < min_diff:
                min_diff = color_diff
                min_dmc = dmc
    else:
        for dmc in dmc_dict:
            dmc_rgb = dmc_dict[dmc]['rgb']

            # pixel rgb
            color1_rgb = sRGBColor(pixel_rgb[0], pixel_rgb[1], pixel_rgb[2])

            # dmc rgb
            color2_rgb = sRGBColor(dmc_rgb[0], dmc_rgb[1], dmc_rgb[2])

            # Convert from RGB to Lab Color Space
            color1_lab = convert_color(color1_rgb, LabColor)

            # Convert from RGB to Lab Color Space
            color2_lab = convert_color(color2_rgb, LabColor)

            # Find the color difference
            delta_e = delta_e_cie2000(color1_lab, color2_lab)

            if delta_e < min_diff:
                min_diff = delta_e
                min_dmc = dmc

    print(min_diff)
    print(dmc_dict[min_dmc]['color'])
    return min_dmc


img = Image.open(img_path)
img_data = np.asarray(img)
# out_data = img_data.copy()
out_data = np.zeros_like(img_data)

del_rows = []
for y in range(0, img_data.shape[0]):
    is_empty_row = True
    for x in range(0, img_data.shape[1]):
        pixel_rgb = img_data[y, x]
        if sum(pixel_rgb[0:2]) == 765 or pixel_rgb[3] == 0:
            pass
        else:
            is_empty_row = False
            break
    if is_empty_row:
        del_rows.append(y)
    else:
        pass

del_cols = []
for x in range(0, img_data.shape[1]):
    is_empty_column = True
    for y in range(0, img_data.shape[0]):
        pixel_rgb = img_data[y, x]
        if sum(pixel_rgb[0:2]) == 765 or pixel_rgb[3] == 0:
            pass
        else:
            is_empty_column = False
            break
    if is_empty_column:
        del_cols.append(x)
    else:
        pass

img_data = np.delete(img_data, del_rows, 1)
img_data = np.delete(img_data, del_cols, 0)
# img_data = np.delete(img_data, np.s_[0:y-1], 0)

workbook = xlsxwriter.Workbook(output_folder + 'out.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:CA', 2.54)
worksheet.set_column('CA:MA', 2.54)

dmc_used_letters = []

# DMC-ify the array
for x in range(0, img_data.shape[0]):
    for y in range(0, img_data.shape[1]):
        pixel_rgb = img_data[x, y]
        if pixel_rgb[3] != 0:
            dmc_key = get_closest_color(pixel_rgb, dmc_map)
            new_rgb = list(dmc_map[dmc_key]['rgb'])
            new_rgb.append(255)
            new_rgb = tuple(new_rgb)

            out_data[x, y] = new_rgb

            cell_format = workbook.add_format()
            cell_format.set_bg_color('#' + dmc_map[dmc_key]['hex'])

            if 'letter' in dmc_map[dmc_key].keys():
                letter = dmc_map[dmc_key]['letter']
            else:
                letter = chars[color_counter % len(chars)]
                dmc_map[dmc_key]['letter'] = letter
                dmc_used_letters.append([letter, dmc_key, dmc_map[dmc_key]['hex']])
                color_counter += 1

            worksheet.write_column(x + 1, y + 1, letter, cell_format)
        else:
            pass

# Numbering Height
height = 0
for y in range(1, img_data.shape[0] + 1):
   height = height + 1
   worksheet.write_column(y, 0, [str(height % 10)])
# Numbering Width
width = 0
for x in range(1, img_data.shape[1] + 1):
    width = width + 1
    worksheet.write_column(height + 1, x, [str(width % 10)])
worksheet.write_column(0, 0, [f'{img_name} [{width} x {height}]'])

# Code and DMC Output
for idx, val in enumerate(dmc_used_letters):
    worksheet.write_column(img_data.shape[0] + 3 + idx, 0, [str(val[0])])
    worksheet.write_column(img_data.shape[0] + 3 + idx, 2, ['Code'])
    worksheet.write_column(img_data.shape[0] + 3 + idx, 4, [str(val[1])])
    worksheet.write_column(img_data.shape[0] + 3 + idx, 6, ['Hex'])
    worksheet.write_column(img_data.shape[0] + 3 + idx, 8, [str(val[2])])

# Insert an image.
# worksheet.insert_image('B5', 'test.png')

workbook.close()

out = Image.fromarray(out_data)
out.save(output_folder + 'test.png')
print("--- %s seconds ---" % (time.time() - start_time))
