# -*-coding: UTF-8 -*-
# Author Shi

from pdf2image import convert_from_path
from pptx import Presentation
import os

file_list = os.listdir('convert')
for file in file_list:
    os.mkdir('temp')
    convert_from_path('convert/' + file, output_folder='temp', fmt='JPG')
    img_list = os.listdir('temp')
    for i in range(0, len(img_list)):
        os.rename('temp/' + img_list[i], 'temp/' + str(i) + ".jpg")
