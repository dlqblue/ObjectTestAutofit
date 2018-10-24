#!/usr/bin/env python3
# -*- coding: utf-8 -*-


data_type_list = [
    '700lux', '100lux', '20lux', '10lux',   # 0-3 TE42
    'resolution',   # 4 T268H
    'dark flash', '5lux flash', '10lux flash', '20lux flash',   # 5-8 Flash TE42
    'Color D65 700lux', 'Color D50 700lux', 'Color TL84 700lux', 'Color A 700lux',   # 9-24 Color Check 24
    'Color D65 100lux', 'Color D50 100lux', 'Color TL84 100lux', 'Color A 100lux',
    'Color D65 20lux', 'Color D50 20lux', 'Color TL84 20lux', 'Color A 20lux',
    'Color D65 10lux', 'Color D50 10lux', 'Color TL84 10lux', 'Color A 10lux',
    'Shading D65 1000lux', 'Shading TL84 1000lux', 'Shading A 1000lux',    # 25-30 DNP
    'Shading D65 20lux', 'Shading TL84 20lux', 'Shading A 20lux',
    'Auto 700lux', 'Auto 100lux', 'Auto 20lux', 'Auto 5lux',   # 31-38 DxO Dead Leaves
    'Touch 700lux', 'Touch 100lux', 'Touch 20lux', 'Touch 5lux',
]

# 西门子星图排列顺序
simens_sort_list = [
    'Star0',
    'Star3',
    'Star7',
    'Star5',
    'Star1',
    'Star4',
    'Star6',
    'Star8',
    'Star2',
    'Star13',
    'Star21',
    'Star14',
    'Star20',
    'Star22',
    'Star12',
    'Star17',
    'Star9',
    'Star16',
    'Star18',
    'Star24',
    'Star10',
    'Star15',
    'Star19',
    'Star23',
    'Star11'
]

# 需要读IE文本的行数
read_ringing_line_number = [188, 189, 190, 191]
read_texture_line_number = [233, 234]
read_noise_line_number = list(range(271, 291))
read_resolution_line_number = list(range(19, 44))
read_flash_AWB_line_number = 18
read_flash_texture_line_number = 234
read_flash_shading_line_number = 242
read_flash_noise_line_number = read_noise_line_number

# 需要读Imatest Color的行数
read_color_line_number = [8, 9, 10] + list(range(85, 109)) + [143]

# 需要读Imatest Shading的行数
read_shading_line_number = [14, 27, 28]

# 对应表格的行数
texture_excel_row = [
    (41, 47),  # 700lux Full, 700lux Low
    (42, 48),  # 100lux Full, 100lux Low
    (43, 49),  # 20lux Full, 20lux Low
    (44, 50)  # 10lux Full, 10lux Low
]
ringing_excel_row = [
    (41, 47, 53, 59),  # 700lux TR_60, 700lux TR_80, 700lux LL_60, 700lux LL_80
    (42, 48, 54, 60),  # 100lux TR_60, 100lux TR_80, 100lux LL_60, 100lux LL_80
    (43, 49, 55, 61),  # 20lux TR_60, 20lux TR_80, 20lux LL_60, 20lux LL_80
    (44, 50, 56, 62)  # 10lux TR_60, 10lux TR_80, 10lux LL_60, 10lux LL_80
]
noise_data_type_dic = {
    0: 'Avg VN(3)',
    1: 'Avg d_L(3)',
    2: 'Avg d_u(3)',
    3: 'Avg d_v(3)'
}
noise_excel_row = [
    (41, 47, 53, 59),  # 700lux Avg VN(3), 700lux Avg d_L(3), 700lux Avg d_u(3), 700lux Avg d_v(3)
    (42, 48, 54, 60),  # 100lux Avg VN(3), 100lux Avg d_L(3), 100lux Avg d_u(3), 100lux Avg d_v(3)
    (43, 49, 55, 61),  # 20lux Avg VN(3), 20lux Avg d_L(3), 20lux Avg d_u(3), 20lux Avg d_v(3)
    (44, 50, 56, 62)  # 10lux Avg VN(3), 10lux Avg d_L(3), 10lux Avg d_u(3), 10lux Avg d_v(3)
]
resolution_excel_row = list(range(64, 89))
flash_excel_row = [39, 40, 41, 42]
flash_data_type_dic = {
    0: 'AWB',
    1: 'Max Attenuation',
    2: 'Visual Noise',
    3: 'Texture Vmtf'
}
color_saturation_excel_row = [
    47, 48, 49, 50,   # 700lux D65, 700lux D50, 700lux TL84, 700lux A
    53, 54, 55, 56,   # 100lux D65, 100lux D50, 100lux TL84, 100lux A
    59, 60, 61, 62,   # 20lux D65, 20lux D50, 20lux TL84, 20lux A
    65, 66, 67, 68    # 10lux D65, 10lux D50, 10lux TL84, 10lux A
]
color_awb_excel_row = [
    42, 43, 44, 45,   # 700lux D65, 700lux D50, 700lux TL84, 700lux A
    49, 50, 51, 52,   # 100lux D65, 100lux D50, 100lux TL84, 100lux A
    56, 57, 58, 59,   # 20lux D65, 20lux D50, 20lux TL84, 20lux A
    63, 64, 65, 66    # 10lux D65, 10lux D50, 10lux TL84, 10lux A
]
color_delta_e_excel_row = [
    list(range(41, 65)),    # D65
    list(range(67, 91)),    # D50
    list(range(93, 117)),   # TL84
    list(range(119, 143))   # A
]
color_data_type_dic = {
    0: 'HSV',
    1: 'E*ab',
    2: 'Saturation'
}

lens_shading_excel_row = [41, 1, 43, 1]   # DNP光箱更换，目前只测试1000， 20两个规格
allow_lens_shading_type = {'Shading D65 1000lux': 0, 'Shading D65 20lux': 2}
color_shading_excel_row = [
    (42, 48),   # 1000lux D65
    (43, 49),   # 1000lux TL84
    (44, 50),   # 1000lux A
    (66, 72),   # 20lux D65
    (67, 73),   # 20lux TL84
    (68, 74)    # 20lux A
]
shading_data_type = {
    0: 'Worst corner level (%)',
    1: 'R/G',
    2: 'B/G'
}

focus_excel_row = [
    list(range(41, 71)),    # 700lux
    list(range(74, 104)),   # 100lux
    list(range(107, 137)),  # 20lux
    list(range(140, 170))   # 5lux
]

focus_data_type = [
    '700lux',
    '100lux',
    '20lux',
    '5lux'
]


def init_global_data():    # 初始化
    global _global_dict
    global _data_dic
    _global_dict = {}
    _data_dic = {}


def set_device_title(key, value):

    _global_dict[key] = value


def get_device_title(key, def_value = None):

    try:
        return _global_dict[key]

    except KeyError:

        return def_value


def set_data_dic(key, value):

    _data_dic[key] = value


def get_data_dic():

    return _data_dic
