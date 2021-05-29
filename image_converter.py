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

# img_path = 'ultraball.png'
img_name = 'Peter_Griffin.png'
dmc_map = 'all'  # all or owned

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
out_data = img_data.copy()

workbook = xlsxwriter.Workbook(output_folder + 'out.xlsx')
worksheet = workbook.add_worksheet()
worksheet.set_column('A:CA', 2.54)
worksheet.set_column('CA:MA', 2.54)

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
                color_counter += 1

            worksheet.write_column(x, y, letter, cell_format)
        else:
            pass


# Insert an image.
# worksheet.insert_image('B5', 'test.png')

workbook.close()

out = Image.fromarray(out_data)
out.save(output_folder + 'test.png')
print("--- %s seconds ---" % (time.time() - start_time))
