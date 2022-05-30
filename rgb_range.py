from PIL import Image
from xlwt import Workbook
import numpy as np
import os

wb = Workbook()

sheet1 = wb.add_sheet('Sheet 1')

sheet1.write(0, 1, 'UPPER RED')
sheet1.write(0, 2, 'LOWER RED')
sheet1.write(0, 3, 'UPPER BLUE')
sheet1.write(0, 4, 'LOWER BLUE')
sheet1.write(0, 5, 'UPPER GREEN')
sheet1.write(0, 6, 'LOWER GREEN')
sheet1.write(0, 7, 'AVERAGE RED')
sheet1.write(0, 8, 'AVERAGE BLUE')
sheet1.write(0, 9, 'AVERAGE GREEN')


def write_into_word(folder, curr_dir):
    os.chdir(folder)
    count = 1
    for imagePath in os.listdir(folder):
        image = Image.open(imagePath)
        # print(imagePath)
        pix_val = list(image.getdata())
        avg_intensity = np.mean(pix_val, axis=0)
        dict_val = {'ured': 0, 'ublue': 0, 'ugreen': 0,
                    'lred': 500, 'lblue': 500, 'lgreen': 500,
                    }
        avg_value = np.max(pix_val, axis=0)
        print(avg_value)
        for coord in pix_val[:]:
            red = coord[0]
            green = coord[1]
            blue = coord[2]
            if dict_val['ured'] < red:
                dict_val['ured'] = red
            if dict_val['ublue'] < blue:
                dict_val['ublue'] = blue
            if dict_val['ugreen'] < green:
                dict_val['ugreen'] = green
            if dict_val['lred'] > red:
                dict_val['lred'] = red
            if dict_val['lblue'] > blue:
                dict_val['lblue'] = blue
            if dict_val['lgreen'] > green:
                dict_val['lgreen'] = green

        sheet1.write(count, 1, dict_val['ured'])
        sheet1.write(count, 2, dict_val['ublue'])
        sheet1.write(count, 3, dict_val['ugreen'])
        sheet1.write(count, 4, dict_val['lred'])
        sheet1.write(count, 5, dict_val['lblue'])
        sheet1.write(count, 6, dict_val['lgreen'])
        sheet1.write(count, 7, avg_intensity[0])
        sheet1.write(count, 8, avg_intensity[1])
        sheet1.write(count, 9, avg_intensity[2])

        count += 1
    os.chdir(curr_dir)
    wb.save('seminar.xls')


folder1 = input('Please enter the folder name: ')
curr_dir = os.getcwd()
folder = os.getcwd() + f'\\{folder1}'
write_into_word(folder, curr_dir)
