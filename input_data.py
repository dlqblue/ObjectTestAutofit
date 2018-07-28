#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import win32com.client as win32
import numpy as np
import data_model
import re


def input_texture_data(wb, obj_data):
    ws = wb.Worksheets('Texture')

    max_column = ws.Range("IV40").End(win32.constants.xlToLeft).Column

    # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
    cell_range_1 = 'IV' + str(data_model.texture_excel_row[obj_data.type_index][0])
    max_cells_column_1 = ws.Range(cell_range_1).End(win32.constants.xlToLeft).Column
    cell_range_2 = 'IV' + str(data_model.texture_excel_row[obj_data.type_index][1])
    max_cells_column_2 = ws.Range(cell_range_2).End(win32.constants.xlToLeft).Column

    allow_write = {max_cells_column_1, max_column}

    if len(allow_write) < 2:

        for x in np.arange(40, 47, 6):
            ws.Cells(int(x), max_column + 1).Value = data_model.get_device_title('device_title')
            ws.Cells(int(x), max_column + 1).Borders.LineStyle = 1

    ws.Cells(data_model.texture_excel_row[obj_data.type_index][0], max_cells_column_1 + 1).Value = obj_data.data_list['texture']['Full DL_cross']
    ws.Cells(data_model.texture_excel_row[obj_data.type_index][0], max_cells_column_1 + 1).Borders.LineStyle = 1
    ws.Cells(data_model.texture_excel_row[obj_data.type_index][1], max_cells_column_2 + 1).Value = obj_data.data_list['texture']['Low DL_cross']
    ws.Cells(data_model.texture_excel_row[obj_data.type_index][1], max_cells_column_2 + 1).Borders.LineStyle = 1

    ws.Columns(max_column + 1).HorizontalAlignment = win32.constants.xlCenter
    ws.Columns(max_column + 1).ColumnWidth = 15.5


def input_ringing_data(wb, obj_data):
    ws = wb.Worksheets('Ringing')

    max_column = ws.Range("IV40").End(win32.constants.xlToLeft).Column

    # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
    cell_range_1 = 'IV' + str(data_model.ringing_excel_row[obj_data.type_index][0])
    max_cells_column_1 = ws.Range(cell_range_1).End(win32.constants.xlToLeft).Column
    cell_range_2 = 'IV' + str(data_model.ringing_excel_row[obj_data.type_index][1])
    max_cells_column_2 = ws.Range(cell_range_2).End(win32.constants.xlToLeft).Column
    cell_range_3 = 'IV' + str(data_model.ringing_excel_row[obj_data.type_index][2])
    max_cells_column_3 = ws.Range(cell_range_3).End(win32.constants.xlToLeft).Column
    cell_range_4 = 'IV' + str(data_model.ringing_excel_row[obj_data.type_index][3])
    max_cells_column_4 = ws.Range(cell_range_4).End(win32.constants.xlToLeft).Column

    allow_write = {max_cells_column_1, max_column}

    if len(allow_write) < 2:

        for x in np.arange(40, 59, 6):
            ws.Cells(int(x), max_column + 1).Value = data_model.get_device_title('device_title')
            ws.Cells(int(x), max_column + 1).Borders.LineStyle = 1

    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][0], max_cells_column_1 + 1).Value = obj_data.data_list['ringing']['Edges_TR Edge60']
    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][1], max_cells_column_2 + 1).Value = obj_data.data_list['ringing']['Edges_TR Edge80']
    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][2], max_cells_column_3 + 1).Value = obj_data.data_list['ringing']['Edges_LL Edge60']
    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][3], max_cells_column_4 + 1).Value = obj_data.data_list['ringing']['Edges_LL Edge80']
    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][0], max_cells_column_1 + 1).Borders.LineStyle = 1
    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][1], max_cells_column_2 + 1).Borders.LineStyle = 1
    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][2], max_cells_column_3 + 1).Borders.LineStyle = 1
    ws.Cells(data_model.ringing_excel_row[obj_data.type_index][3], max_cells_column_4 + 1).Borders.LineStyle = 1

    ws.Columns(max_column + 1).HorizontalAlignment = win32.constants.xlCenter
    ws.Columns(max_column + 1).ColumnWidth = 15.5


def input_noise_data(wb, obj_data):
    ws = wb.Worksheets('Noise')

    max_column = ws.Range("IV40").End(win32.constants.xlToLeft).Column

    # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
    cell_range_1 = 'IV' + str(data_model.noise_excel_row[obj_data.type_index][0])
    max_cells_column_1 = ws.Range(cell_range_1).End(win32.constants.xlToLeft).Column
    cell_range_2 = 'IV' + str(data_model.noise_excel_row[obj_data.type_index][1])
    max_cells_column_2 = ws.Range(cell_range_2).End(win32.constants.xlToLeft).Column
    cell_range_3 = 'IV' + str(data_model.noise_excel_row[obj_data.type_index][2])
    max_cells_column_3 = ws.Range(cell_range_3).End(win32.constants.xlToLeft).Column
    cell_range_4 = 'IV' + str(data_model.noise_excel_row[obj_data.type_index][3])
    max_cells_column_4 = ws.Range(cell_range_4).End(win32.constants.xlToLeft).Column

    allow_write = {max_cells_column_1, max_column}

    if len(allow_write) < 2:

        for x in np.arange(40, 59, 6):
            ws.Cells(int(x), max_column + 1).Value = data_model.get_device_title('device_title')
            ws.Cells(int(x), max_column + 1).Borders.LineStyle = 1

    ws.Cells(data_model.noise_excel_row[obj_data.type_index][0], max_cells_column_1 + 1).Value = obj_data.data_list['noise'][data_model.noise_data_type_dic[0]]
    ws.Cells(data_model.noise_excel_row[obj_data.type_index][1], max_cells_column_2 + 1).Value = obj_data.data_list['noise'][data_model.noise_data_type_dic[1]]
    ws.Cells(data_model.noise_excel_row[obj_data.type_index][2], max_cells_column_3 + 1).Value = obj_data.data_list['noise'][data_model.noise_data_type_dic[2]]
    ws.Cells(data_model.noise_excel_row[obj_data.type_index][3], max_cells_column_4 + 1).Value = obj_data.data_list['noise'][data_model.noise_data_type_dic[3]]
    ws.Cells(data_model.noise_excel_row[obj_data.type_index][0], max_cells_column_1 + 1).Borders.LineStyle = 1
    ws.Cells(data_model.noise_excel_row[obj_data.type_index][1], max_cells_column_2 + 1).Borders.LineStyle = 1
    ws.Cells(data_model.noise_excel_row[obj_data.type_index][2], max_cells_column_3 + 1).Borders.LineStyle = 1
    ws.Cells(data_model.noise_excel_row[obj_data.type_index][3], max_cells_column_4 + 1).Borders.LineStyle = 1

    ws.Columns(max_column + 1).HorizontalAlignment = win32.constants.xlCenter
    ws.Columns(max_column + 1).ColumnWidth = 15.5


def input_resolution_data(wb, obj_data):
    ws = wb.Worksheets('Resolution-HW')

    max_column = ws.Range("IV62").End(win32.constants.xlToLeft).Column

    ws.Cells(62, max_column + 2).Value = data_model.get_device_title('device_title')
    ws.Range(ws.Cells(62, max_column + 2), ws.Cells(62, max_column + 3)).MergeCells = True
    ws.Range(ws.Cells(62, max_column + 2), ws.Cells(62, max_column + 3)).HorizontalAlignment = win32.constants.xlCenter
    ws.Cells(63, max_column + 2).Value = 'Result'
    ws.Cells(63, max_column + 3).Value = 'Cor/Cen'

    # ws.Cells(2, 2).Font.Color = -16776961

    for x in range(25):
        ws.Cells(data_model.resolution_excel_row[x], max_column + 2).Value = obj_data.data_list['resolution'][x][0]
        ws.Cells(data_model.resolution_excel_row[x], max_column + 3).Value = obj_data.data_list['resolution'][x][1]

    ws.Range(ws.Cells(62, max_column + 2), ws.Cells(data_model.resolution_excel_row[24], max_column + 3)).Borders.LineStyle = 1
    ws.Columns(max_column + 2).HorizontalAlignment = win32.constants.xlCenter
    ws.Columns(max_column + 3).HorizontalAlignment = win32.constants.xlCenter


def input_flash_data(wb, obj_data):
    ws = wb.Worksheets('Flash-SW')

    max_column = ws.Range("IV38").End(win32.constants.xlToLeft).Column

    ws.Cells(38, max_column + 1).Value = data_model.get_device_title('device_title') + '\n' +obj_data.data_type

    ws.Cells(data_model.flash_excel_row[0], max_column + 1).Value = obj_data.data_list['flash_awb']['WB [CIE-C]']
    ws.Cells(data_model.flash_excel_row[1], max_column + 1).Value = obj_data.data_list['flash_shading']['Shading [%]'] + '%'
    ws.Cells(data_model.flash_excel_row[2], max_column + 1).Value = obj_data.data_list['flash_texture']['Full DL_cross']
    ws.Cells(data_model.flash_excel_row[3], max_column + 1).Value = obj_data.data_list['flash_noise']['Avg VN(3)']

    ws.Range(ws.Cells(38, max_column + 1), ws.Cells(data_model.flash_excel_row[3], max_column + 1)).Borders.LineStyle = 1
    ws.Columns(max_column + 1).HorizontalAlignment = win32.constants.xlCenter
    ws.Columns(max_column + 1).ColumnWidth = 15


def input_color_saturation_data(wb, obj_data):

    ws = wb.Worksheets('Color-Saturation')
    max_column = ws.Range("IV46").End(win32.constants.xlToLeft).Column

    # 为了data_type_index和color_saturation_row可以对应起来
    row = obj_data.type_index - 9

    # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
    cell_range = 'IV' + str(data_model.color_saturation_excel_row[row])
    max_cells_column = ws.Range(cell_range).End(win32.constants.xlToLeft).Column

    if max_column == max_cells_column:

        for x in np.arange(46, 65, 6):

            ws.Cells(int(x), max_column + 1).Value = data_model.get_device_title('device_title')
            ws.Cells(int(x), max_column + 1).Borders.LineStyle = 1

    ws.Cells(data_model.color_saturation_excel_row[row], max_cells_column + 1).Value = obj_data.data_list['color'][data_model.color_data_type_dic[2]] + '%'
    ws.Cells(data_model.color_saturation_excel_row[row], max_cells_column + 1).Borders.LineStyle = 1

    ws.Columns(max_cells_column + 1).HorizontalAlignment = win32.constants.xlCenter
    ws.Columns(max_cells_column + 1).ColumnWidth = 15.5


def input_color_awb_data(wb, obj_data):

    ws = wb.Worksheets('Color-AWB')
    max_column = ws.Range("IV40").End(win32.constants.xlToLeft).Column

    row = obj_data.type_index - 9

    # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
    cell_range = 'IV' + str(data_model.color_awb_excel_row[row])
    max_cells_column = ws.Range(cell_range).End(win32.constants.xlToLeft).Column

    if max_column + 2 == max_cells_column:

        for x in np.arange(40, 62, 7):

            ws.Cells(int(x), max_column + 3).Value = data_model.get_device_title('device_title')
            ws.Range(ws.Cells(int(x), max_column + 3), ws.Cells(int(x), max_column + 5)).MergeCells = True
            ws.Cells(int(x) + 1, max_column + 3).Value = 'Block21'
            ws.Cells(int(x) + 1, max_column + 4).Value = 'Block22'
            ws.Cells(int(x) + 1, max_column + 5).Value = 'Block23'

    for x in range(3):
        ws.Cells(data_model.color_awb_excel_row[row], max_cells_column + 1 + x).Value = obj_data.data_list['color'][str(x + 21) + 'HSV']
        ws.Cells(data_model.color_awb_excel_row[row], max_cells_column + 1 + x).HorizontalAlignment = win32.constants.xlCenter

    if max_column + 2 == max_cells_column:

        ws.Range(ws.Cells(40, max_column + 3), ws.Cells(data_model.color_awb_excel_row[-1], max_column + 5)).Borders.LineStyle = 1
        ws.Range(ws.Cells(40, max_column + 3), ws.Cells(data_model.color_awb_excel_row[-1], max_column + 5)).HorizontalAlignment = win32.constants.xlCenter


def input_color_delta_e_data(wb, obj_data):

    # Delta E只读700lux
    allow_data_type = [
        data_model.data_type_list[9],
        data_model.data_type_list[10],
        data_model.data_type_list[11],
        data_model.data_type_list[12]
    ]

    if obj_data.data_type in allow_data_type:

        ws = wb.Worksheets('Color-Accuracy')
        max_column = ws.Range("IV40").End(win32.constants.xlToLeft).Column

        # 为了data_type_index和color_saturation_row可以对应起来
        row = obj_data.type_index - 9

        # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
        cell_range = 'IV' + str(data_model.color_delta_e_excel_row[row][1])
        max_cells_column = ws.Range(cell_range).End(win32.constants.xlToLeft).Column

        if max_column == max_cells_column:

            for x in np.arange(40, 119, 26):

                ws.Cells(int(x), max_column + 1).Value = data_model.get_device_title('device_title')
                ws.Cells(int(x), max_column + 1).Borders.LineStyle = 1

        for x in range(24):
            ws.Cells(data_model.color_delta_e_excel_row[row][x], max_cells_column + 1).Value = obj_data.data_list['color'][data_model.color_data_type_dic[1] + str(x + 1)]
            ws.Cells(data_model.color_delta_e_excel_row[row][x], max_cells_column + 1).Borders.LineStyle = 1

        ws.Columns(max_cells_column + 1).HorizontalAlignment = win32.constants.xlCenter
        ws.Columns(max_cells_column + 1).ColumnWidth = 15.5


def input_lens_shading(wb, obj_data):

    ws = wb.Worksheets('LensShading')
    max_cloumn = ws.Range("IV40").End(win32.constants.xlToLeft).Column

    # LensShading只读D65，index 25, 28
    if obj_data.data_type in data_model.allow_lens_shading_type:

        # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
        cell_range = 'IV' + str(data_model.lens_shading_excel_row[data_model.allow_lens_shading_type[obj_data.data_type]])
        max_cells_column = ws.Range(cell_range).End(win32.constants.xlToLeft).Column

        if max_cells_column == max_cloumn:
            ws.Cells(40, max_cloumn + 1).Value = data_model.get_device_title('device_title')
            ws.Cells(40, max_cloumn + 1).Borders.LineStyle = 1

        ws.Cells(data_model.lens_shading_excel_row[data_model.allow_lens_shading_type[obj_data.data_type]], max_cells_column + 1).Value = obj_data.data_list['shading'][data_model.shading_data_type[0]] + '%'
        ws.Cells(data_model.lens_shading_excel_row[data_model.allow_lens_shading_type[obj_data.data_type]], max_cells_column + 1).Borders.LineStyle = 1

        ws.Columns(max_cells_column + 1).HorizontalAlignment = win32.constants.xlCenter
        ws.Columns(max_cells_column + 1).ColumnWidth = 15.5


def input_color_shading(wb, obj_data):

    ws = wb.Worksheets('ColorShading')
    max_column = ws.Range("IV40").End(win32.constants.xlToLeft).Column

    row = obj_data.type_index - 25

    # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
    cell_range = 'IV' + str(data_model.color_shading_excel_row[row][0])
    max_cells_column = ws.Range(cell_range).End(win32.constants.xlToLeft).Column

    if max_column + 2 == max_cells_column:

        for x in np.arange(40, 83, 6):

            ws.Cells(int(x), max_column + 3).Value = data_model.get_device_title('device_title')
            ws.Range(ws.Cells(int(x), max_column + 3), ws.Cells(int(x), max_column + 5)).MergeCells = True
            ws.Cells(int(x) + 1, max_column + 3).Value = 'RG_Max'
            ws.Cells(int(x) + 1, max_column + 4).Value = 'BG_Min'
            ws.Cells(int(x) + 1, max_column + 5).Value = 'Max colored vignetting'

            ws.Range(ws.Cells(int(x), max_column + 3), ws.Cells(int(x), max_column + 5)).Borders.LineStyle = 1
            ws.Cells(int(x) + 1, max_column + 3).Borders.LineStyle = 1
            ws.Cells(int(x) + 1, max_column + 4).Borders.LineStyle = 1
            ws.Cells(int(x) + 1, max_column + 5).Borders.LineStyle = 1

            ws.Cells(int(x), max_column + 3).HorizontalAlignment = win32.constants.xlCenter
            ws.Cells(int(x) + 1, max_column + 3).HorizontalAlignment = win32.constants.xlCenter
            ws.Cells(int(x) + 1, max_column + 4).HorizontalAlignment = win32.constants.xlCenter
            ws.Cells(int(x) + 1, max_column + 5).HorizontalAlignment = win32.constants.xlCenter
            ws.Columns(max_column + 5).ColumnWidth = 25.5

    for x in range(2):

        max_result = float(obj_data.data_list['shading'][data_model.shading_data_type[1] + 'Maximum'])
        min_result = float(obj_data.data_list['shading'][data_model.shading_data_type[1] + 'Minimum'])
        vignetting = str((max_result - min_result ) / max_result) + '%'
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 1).Value = max_result
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 2).Value = min_result
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 3).Value = vignetting
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 1).Borders.LineStyle = 1
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 2).Borders.LineStyle = 1
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 3).Borders.LineStyle = 1
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 1).HorizontalAlignment = win32.constants.xlCenter
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 2).HorizontalAlignment = win32.constants.xlCenter
        ws.Cells(data_model.color_shading_excel_row[row][x], max_cells_column + 3).HorizontalAlignment = win32.constants.xlCenter


def input_focus_data(wb, data):

    for x in data:

        for key in x:
            data_type = re.split(r"[ ]", key)
            data_index = data_model.focuc_data_type.index(data_type[-1])
            ws = wb.Worksheets('Focus-' + data_type[0])

            # 写入title之后有效列会变，影响后续数据输入，所以要单独算每一行的有效列
            cell_range = 'IV' + str(data_model.focus_excel_row[data_index][0])
            max_cells_column = ws.Range(cell_range).End(win32.constants.xlToLeft).Column

            ws.Cells(int(data_model.focus_excel_row[data_index][0]) - 1, max_cells_column + 1).Value = data_model.get_device_title('device_title')

            for y in range(30):
                ws.Cells(data_model.focus_excel_row[data_index][y], max_cells_column + 1).Value = '{:.3f}%'.format(x[key][y])

            ws.Cells(data_model.focus_excel_row[data_index][29] + 1, max_cells_column + 1).Value = '{:.2f}%'.format(np.std(x[key], ddof = 1) / np.mean(x[key]) * 100)
            ws.Cells(data_model.focus_excel_row[data_index][29] + 1, max_cells_column + 1).Interior.Color = rgb_to_int((255, 255, 0))
            ws.Range(ws.Cells(int(data_model.focus_excel_row[data_index][0]) - 1, max_cells_column + 1), ws.Cells(data_model.focus_excel_row[data_index][29] + 1, max_cells_column + 1)).HorizontalAlignment = win32.constants.xlCenter
            ws.Range(ws.Cells(int(data_model.focus_excel_row[data_index][0]) - 1, max_cells_column + 1),
                     ws.Cells(data_model.focus_excel_row[data_index][29] + 1, max_cells_column + 1)).Borders.LineStyle = 1
            ws.Columns(max_cells_column + 1).ColumnWidth = 15.5


def rgb_to_int(rgb):

    color_int = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)

    return color_int
