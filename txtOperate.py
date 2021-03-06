#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import win32com.client as win32
import re
import os
import input_data
import data_model
import read_data

# 导入Excel库，用这个方法导入Excel库，可以直接读到win32.constants.xlCenter
excel = win32.gencache.EnsureDispatch('Excel.Application')
# 关闭警告对话框
excel.DisplayAlerts = False

focus_data_list = []  # 对焦的数据集


class ObjDataModel(object):

    def __init__(self, data_type, type_index=-1, data_list=None):
        if data_list is None:
            data_list = {}
        self.data_list = data_list
        self.data_type = data_type
        self.type_index = type_index


def read_txt(path):
    # path_example = \\DXO Object test database\Database NEW\17819\20180704_P\10lux_IMG_20180103_064818_42.txt
    # 解析文件名称，判定文件类型
    file_type_list = re.split(r"[\\,_,.]", path)

    # 解析文件名跟字典匹配

    file_type = [x for x in file_type_list[-10:] if x in data_model.data_type_list][0]
    type_index = data_model.data_type_list.index(file_type)
    print('-------------------------------data here-------------------------------')
    print('\nFile: ' + path + '\n')

    if type_index <= data_model.data_type_index[0]:  # TE42

        print('This is ' + file_type + ' TE42!\n\n')
        # 生成对象赋予数据类型
        obj_data = ObjDataModel(file_type, type_index)

        with open(path, 'r') as readFile:

            read_data.read_ringing(readFile)  # 读ringing

            read_data.read_texture(readFile)  # 读texture

            read_data.read_noise(readFile)  # 读noise

            obj_data.data_list = data_model.get_data_dic()

    elif type_index == data_model.data_type_index[1]:  # T268H

        print('This is ' + file_type + '!\n')

        obj_data = ObjDataModel(file_type, type_index)

        with open(path, 'r') as readFile:

            read_data.read_resolution(readFile)

            obj_data.data_list = data_model.get_data_dic()

    elif data_model.data_type_index[2] >= type_index > data_model.data_type_index[1]:  # Flash TE42

        print('This is ' + file_type + ' TE42 !\n')

        obj_data = ObjDataModel(file_type, type_index)

        with open(path, 'r') as readFile:

            read_data.read_flash_awb(readFile)

            read_data.read_flash_texture(readFile)

            read_data.read_flash_shading(readFile)

            read_data.read_flash_noise(readFile)

            obj_data.data_list = data_model.get_data_dic()

    elif data_model.data_type_index[3] >= type_index > data_model.data_type_index[2]:  # Color

        print('This is ' + file_type + ' !\n')

        obj_data = ObjDataModel(file_type, type_index)

        with open(path, 'r') as csv_file:

            read_data.read_color(csv_file)

            obj_data.data_list = data_model.get_data_dic()

    elif data_model.data_type_index[4] >= type_index > data_model.data_type_index[3]:  # Shading

        print('This is ' + file_type + ' !\n')

        obj_data = ObjDataModel(file_type, type_index)

        with open(path, 'r') as csv_file:

            read_data.read_shading(csv_file)

            obj_data.data_list = data_model.get_data_dic()

    elif data_model.data_type_index[5] >= type_index > data_model.data_type_index[4]:

        print('This is ' + file_type + ' focus !\n')

    else:

        print('\nNo TYPE！')

    print(obj_data.data_list)
    print('\n\n')

    return obj_data


def read_focus_file(path_dic, path):
    for key in path_dic:
        data_type = re.split(r"[\\]", key)[-1]
        one_type_data_dic = {}

        if os.path.splitext(path_dic[key][0])[1] == '.xls':
            one_type_data_dic = read_data.read_focus_dxo(data_type, path_dic[key])
        elif os.path.splitext(path_dic[key][0])[1] == '.txt':
            one_type_data_dic = read_data.read_focus_ie(data_type, path_dic[key])
        focus_data_list.append(one_type_data_dic)

    wb = excel.Workbooks.Open(path)

    input_data.input_focus_data(wb, focus_data_list)

    wb.SaveAs(path)
    excel.Application.Quit()


def input_excel(obj_data, path):
    wb = excel.Workbooks.Open(path)

    try:
        if obj_data.type_index <= data_model.data_type_index[0]:

            input_data.input_texture_data(wb, obj_data)
            input_data.input_ringing_data(wb, obj_data)
            input_data.input_noise_data(wb, obj_data)

        elif obj_data.type_index == data_model.data_type_index[1]:

            input_data.input_resolution_data(wb, obj_data)

        elif data_model.data_type_index[2] >= obj_data.type_index > data_model.data_type_index[1]:

            input_data.input_flash_data(wb, obj_data)

        elif data_model.data_type_index[3] >= obj_data.type_index > data_model.data_type_index[2]:

            input_data.input_color_saturation_data(wb, obj_data)
            input_data.input_color_awb_data(wb, obj_data)
            input_data.input_color_delta_e_data(wb, obj_data)

        elif data_model.data_type_index[4] >= obj_data.type_index > data_model.data_type_index[3]:

            input_data.input_lens_shading(wb, obj_data)
            input_data.input_color_shading(wb, obj_data)

    except Exception:
        wb.SaveAs(path)
        excel.Application.Quit()

    finally:
        wb.SaveAs(path)


if __name__ == '__main__':

    data_model.init_global_data()

    txt_path = []

    focus_path_dic = {}

    for x in os.listdir(os.getcwd()):
        if os.path.splitext(x)[1] == '.txt' or os.path.splitext(x)[1] == '.csv':
            txt_path.append(os.path.join(os.getcwd(), x))

    # print('\nPlease input the number of files: \n')
    # num = int(input())
    #
    #
    # for x in range(num):
    #     print('\nPlease input text path: \n')
    #     txt_path.append(input()[1:-1])
    #
    # txt_path = txt_path[1:-1]
    # 文本输入路径

    print('\nPlease excel path: \n')
    excel_path = input()
    excel_path = excel_path[1:-1]
    # excel输入路径

    print('\nPlease input title: \n')

    # device_title = input()
    data_model.set_device_title('device_title', input())

    if os.path.isdir('./Focus'):

        for root, dirs, files in os.walk('.\\Focus'):
            focus_path = []
            if root != '.\\Focus':  # 排除根目录影响
                for file in files:
                    if os.path.splitext(file)[1] == '.txt' or os.path.splitext(file)[1] == '.xls':
                        focus_path.append(os.path.join(os.path.dirname(os.path.abspath(file)), root[2:], file))
                focus_path_dic[root] = focus_path
        print('Focus data file detected, reading...\n')
        read_focus_file(focus_path_dic, excel_path)

    for x in range(len(txt_path)):
        data = read_txt(txt_path[x])
        input_excel(data, excel_path)

    print('\n')
    os.system('pause')
