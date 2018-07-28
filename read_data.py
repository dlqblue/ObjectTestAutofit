#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import data_model
import win32com.client as win32
import txtOperate


def read_ringing(read_file):
    # Ringing
    # 187 # Sub	Edge	(Y)MTF [%] 	vMTF Set 1 	vMTF Set 2	vMTF Set 3	CPIQ_SmallPrint	CPIQ_LargePrint	CPIQ_Monitor
    # 188 #     Edges_TR	Edge60	0.777	0.974	1.149	1.164	0.767	0.751	0.980	1.018	0.955
    # 189 #     Edges_TR	Edge80	0.851	1.001	1.105	1.122	0.736	0.738	0.964	1.004	0.956
    # 190 #     Edges_LL	Edge60	0.771	0.978	1.148	1.167	0.767	0.748	0.985	1.023	0.960
    # 191 #     Edges_LL	Edge80	0.759	0.913	1.068	1.066	0.706	0.729	0.896	0.933	0.881

    read_line = read_file.readline()
    data_list_dic = {}

    count_line = 0
    n = 1 + 1

    while read_line:

        if data_model.read_ringing_line_number[count_line] == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)

            print(data_string[0], data_string[1] + ' : ' + data_string[5])
            data_list_dic[data_string[0] + ' ' + data_string[1]] = data_string[5]

            count_line += 1
            n += 1

            if n > data_model.read_ringing_line_number[-1]:
                break
        else:

            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('ringing', data_list_dic)


def read_texture(read_file):
    # Texture
    # 232 # Sub	Patch	(Y)MTF [%] 	vMTF Set 1 	vMTF Set 2	vMTF Set 3	CPIQ_SmallPrint	CPIQ_LargePrint	CPIQ_Monitor
    # 233 #      Low	 DL_cross	0.289	0.405	0.814	0.702	0.521	0.796	0.481	0.494	0.432
    # 234 #      Full    DL_cross	0.351	0.497	0.896	0.809	0.588	0.795	0.582	0.599	0.528

    read_line = read_file.readline()
    data_list_dic = {}
    count_line = 0
    n = 192 + 1  # 接着ringing后面继续读

    while read_line:

        if data_model.read_texture_line_number[count_line] == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)

            print(data_string[0][1:], data_string[1][1:] + ' : ' + data_string[5])
            data_list_dic[data_string[0][1:] + data_string[1]] = data_string[5]

            count_line += 1
            n += 1

            if n > data_model.read_texture_line_number[-1]:
                break
        else:

            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('texture', data_list_dic)


def read_noise(read_file):
    # Nosie
    # 270  Patch 	 VN (1)  	 VN (2) 	 VN (3)	 d_L (1)  	 d_L (2) 	 dL (3)	 d_u (1)  	 d_u (2) 	 d_u (3)
    # d_v (1)  	 d_v (2) 	 d_v (3)
    # 271     OECF-20 	 0.535 	 0.185 	 0.242	 0.483 	 0.170 	 0.222	 0.048 	 0.016 	 0.021	 0.033 	 0.005 	 0.007
    # 272     OECF-19 	 0.777 	 0.525 	 0.541	 0.626 	 0.394 	 0.405	 0.142 	 0.124 	 0.129	 0.094 	 0.077 	 0.082
    # 273     OECF-18 	 1.186 	 1.030 	 1.009	 1.071 	 0.932 	 0.910	 0.114 	 0.098 	 0.099	 0.055 	 0.043 	 0.045
    # 274     OECF-17 	 0.638 	 0.268 	 0.308	 0.592 	 0.248 	 0.284	 0.045 	 0.019 	 0.022	 0.026 	 0.013 	 0.014
    # 275     OECF-16 	 1.539 	 0.709 	 0.803	 1.404 	 0.643 	 0.726	 0.134 	 0.064 	 0.074	 0.065 	 0.037 	 0.041
    # 276     OECF-15 	 1.731 	 0.587 	 0.776	 1.594 	 0.541 	 0.715	 0.140 	 0.047 	 0.063	 0.057 	 0.017 	 0.023
    # 277     OECF-14 	 2.874 	 1.102 	 1.362	 2.651 	 1.018 	 1.258	 0.235 	 0.091 	 0.112	 0.070 	 0.020 	 0.025
    # 278     OECF-13 	 2.932 	 1.092 	 1.371	 2.599 	 0.862 	 1.119	 0.339 	 0.236 	 0.257	 0.138 	 0.091 	 0.104
    # 279     OECF-12 	 2.755 	 0.998 	 1.287	 2.408 	 0.757 	 1.020	 0.346 	 0.247 	 0.270	 0.160 	 0.095 	 0.111
    # 280     OECF-11 	 3.308 	 1.433 	 1.685	 3.066 	 1.325 	 1.558	 0.265 	 0.118 	 0.138	 0.050 	 0.024 	 0.030
    # 281     OECF-10 	 3.299 	 1.554 	 1.731	 3.029 	 1.429 	 1.589	 0.241 	 0.123 	 0.135	 0.198 	 0.066 	 0.084
    # 282     OECF-9 	 2.834 	 1.187 	 1.400	 2.586 	 1.077 	 1.269	 0.231 	 0.109 	 0.127	 0.160 	 0.052 	 0.072
    # 283     OECF-8 	 2.710 	 1.478 	 1.575	 2.443 	 1.319 	 1.402	 0.205 	 0.124 	 0.129	 0.284 	 0.164 	 0.197
    # 284     OECF-7 	 1.918 	 0.803 	 0.972	 1.614 	 0.582 	 0.726	 0.190 	 0.171 	 0.180	 0.439 	 0.232 	 0.288
    # 285     OECF-6 	 1.733 	 1.010 	 1.099	 1.293 	 0.632 	 0.698	 0.233 	 0.237 	 0.242	 0.748 	 0.541 	 0.605
    # 286     OECF-5 	 1.369 	 0.483 	 0.626	 1.274 	 0.449 	 0.582	 0.102 	 0.036 	 0.047	 0.027 	 0.009 	 0.012
    # 287     OECF-4 	 1.752 	 0.837 	 0.981	 1.298 	 0.476 	 0.591	 0.405 	 0.341 	 0.362	 0.338 	 0.219 	 0.252
    # 288     OECF-3 	 1.343 	 0.486 	 0.608	 1.234 	 0.446 	 0.558	 0.116 	 0.043 	 0.054	 0.031 	 0.011 	 0.015
    # 289     OECF-2 	 1.606 	 0.793 	 0.911	 1.246 	 0.523 	 0.611	 0.367 	 0.280 	 0.310	 0.147 	 0.097 	 0.112
    # 290     OECF-1 	 1.250 	 0.481 	 0.575	 1.128 	 0.428 	 0.510	 0.109 	 0.052 	 0.061	 0.091 	 0.030 	 0.042

    read_line = read_file.readline()
    count_line = 0
    n = 235 + 1     # 接着texture后面继续读

    data_list = [[0 for col in range(4)] for row in range(20)]

    while read_line:

        if data_model.read_noise_line_number[count_line] == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)

            print(data_string[0] + ':' + data_string[3], data_string[6], data_string[9], data_string[12][:-1])

            data_list[count_line][0] = data_string[3][1:]
            data_list[count_line][1] = data_string[6][1:]
            data_list[count_line][2] = data_string[9][1:]
            data_list[count_line][3] = data_string[12][1:-1]

            count_line += 1
            n += 1

            if n > data_model.read_noise_line_number[-1]:
                break

        else:
            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('noise', cal_noise(data_list))


def cal_noise(data_list):

    data_list_dic = {}

    for x in range(4):
        noise = 0
        for y in range(20):
            noise += float(data_list[y][x])
        data_list_dic[data_model.noise_data_type_dic[x]] = noise / 20

    return data_list_dic


def read_resolution(read_file):

    # 17   Siemens	Lim.Resolution [LP/PH]  10%
    # 18   Sub	Mean(Y)	Segment 1(Y)    Segment 2(Y)	Segment 3(Y)	Segment 4(Y)	Segment 5(Y)	Segment 6(Y)	Segment 7(Y)	Segment 8(Y)
    # 19   Star0	1750.93	1615.27	1748.81	3117.94	1793.50	1618.69	1716.02	1839.05	1700.17
    # 20   Star1	1572.32	1533.48	1556.32	1576.74	1593.55	1546.56	1622.40	1575.07	1588.45
    # 21   Star2	1618.41	1681.23	1800.21	1511.41	1387.57	1665.82	1682.52	1483.10	1382.04
    # 22   Star3	1950.25	1721.67	1841.23	-1728.00	1797.50	2540.33	1795.87	2666.16	1812.75
    # 23   Star4	1649.17	1719.62	1366.38	1527.39	1772.22	1690.75	1370.21	1540.57	1753.37
    # 24   Star5	1591.16	1544.99	1647.49	1614.56	1605.18	1543.29	1619.15	1589.33	1620.56
    # 25   Star6	1550.24	1626.25	1682.98	1485.15	1330.27	1615.66	1773.43	1500.79	1335.98
    # 26   Star7	2553.06	2654.26	1814.71	-1728.00	1786.56	2581.75	1840.16	-1728.00	1811.61
    # 27   Star8	1565.08	1546.96	1350.62	1508.97	1716.62	1536.95	1253.23	1494.46	1796.96
    # 28   Star9	1553.37	1540.63	1594.18	1644.19	1509.54	1447.41	1529.19	1633.16	1533.75
    # 29   Star10	1679.80	1769.86	1764.42	1633.98	1575.09	1742.81	1672.66	1586.72	1613.60
    # 30   Star11	1441.49	1273.75	1604.92	1310.70	1174.41	1368.67	1654.83	1399.14	1177.46
    # 31   Star12	1532.27	1426.32	1603.94	1693.73	1466.06	1360.53	1554.00	1759.49	1444.18
    # 32   Star13	1528.95	979.89	1400.66	1818.18	1433.56	1054.26	1427.51	1780.19	1399.77
    # 33   Star14	1586.60	1453.04	1388.23	1788.08	1726.23	1487.07	1378.46	1671.48	1585.01
    # 34   Star15	1527.78	1562.65	1465.86	1467.83	1611.36	1525.64	1464.82	1491.32	1593.13
    # 35   Star16	1702.60	2625.76	1620.90	1438.50	1702.59	2301.32	1600.75	1449.37	1757.16
    # 36   Star17	1516.74	1480.06	1531.53	1440.41	1545.78	1567.85	1578.22	1443.30	1521.91
    # 37   Star18	1653.87	2587.54	1648.49	1425.51	1521.10	1902.85	1678.69	1421.87	1564.43
    # 38   Star19	1341.31	1399.55	1444.47	1251.16	1242.31	1363.26	1441.80	1228.44	1259.26
    # 39   Star20	1435.45	1371.88	1470.87	1467.51	1221.02	1405.59	1613.53	1538.55	1251.20
    # 40   Star21	1459.41	896.56	1280.59	1709.65	1245.58	925.19	1377.37	1791.63	1318.31
    # 41   Star22	1550.40	1400.49	1451.50	1641.00	1507.23	1402.38	1494.80	1697.68	1692.35
    # 42   Star23	1594.59	1619.08	1483.36	1547.21	1654.67	1662.50	1488.11	1554.26	1664.24
    # 43   Star24	1672.55	2005.30	1648.24	1588.74	1636.68	1680.75	1596.75	1646.30	1734.21

    read_line = read_file.readline()
    count_line = 0
    n = 1 + 1

    data_list = {}

    while read_line:

        if data_model.read_resolution_line_number[count_line] == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)

            print(data_string[0] + ' ' + data_string[1])

            data_list[data_string[0]] = data_string[1]

            count_line += 1
            n += 1

            if n > data_model.read_resolution_line_number[-1]:
                break
        else:
            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('resolution', sort_resolution(data_list))


def sort_resolution(data_list):

    # 0F		            Star0
    # O.2F	Top	            Star3
    # 	    Bottom	        Star7
    # 0.34F	Left	        Star5
    # 	    Right	        Star1
    # 0.4F	Top-Left	    Star4
    # 	    Bottom-Left	    Star6
    # 	    Bottom-Right	Star8
    # 	    Top-Right	    Star2
    # 0.41F	Top	            Star13
    # 	    Bottom	        Star21
    # 0.53F	Top-Left	    Star14
    # 	    Bottom-Left	    Star20
    # 	    Bottom-Right	Star22
    # 	    Top-Right	    Star12
    # 0.67F	Left	        Star17
    # 	    Right	        Star9
    # 0.70F	Top-Left	    Star16
    # 	    Bottom-Left	    Star18
    # 	    Bottom-Right	Star24
    # 	    Top-Right	    Star10
    # 0.79F	Top-Left	    Star15
    # 	    Bottom-Left	    Star19
    # 	    Bottom-Right	Star23
    # 	    Top-Right	    Star11

    sort_list = [[0 for col in range(2)] for row in range(25)]

    for x in range(25):

        sort_list[x][0] = float(data_list[data_model.simens_sort_list[x]])

    for x in range(25):

        sort_list[x][1] = '{:.2f}%'.format(sort_list[x][0]/sort_list[0][0]*100)

    return sort_list


def read_flash_awb(read_file):

    # 18    SNR_total [dB]	28.88	d_L (1)	1.714	2.698	DR_total [D]	2.58	WB [CIE-C]	1.781	ISO_S/N10	250

    read_line = read_file.readline()
    n = 1 + 1
    data_list_dic = {}

    while read_file:

        if data_model.read_flash_AWB_line_number == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)
            print(data_string[-4] + ' : ' + data_string[-3])
            data_list_dic[data_string[-4]] = data_string[-3]

            break

        else:

            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('flash_awb', data_list_dic)


def read_flash_texture(read_file):

    # Texture
    # 234 #      Full    DL_cross	0.351	0.497	0.896	0.809	0.588	0.795	0.582	0.599	0.528

    read_line = read_file.readline()
    data_list_dic = {}
    n = 19 + 1  # 接着flash_awb后面继续读

    while read_line:

        if data_model.read_flash_texture_line_number == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)

            print(data_string[0][1:], data_string[1][1:] + ' : ' + data_string[5])
            data_list_dic[data_string[0][1:] + data_string[1]] = data_string[5]

            break

        else:

            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('flash_texture', data_list_dic)


def read_flash_shading(read_file):

    # 242   Shading [%]	51.5	CIE C	2.1	VN Set1	4.4	Optical Center vertical [pixel]	45.0

    read_line = read_file.readline()
    data_list_dic = {}
    n = 235 + 1  # 接着flash_texture后面继续读

    while read_line:

        if data_model.read_flash_shading_line_number == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)

            print(data_string[0] + ' : ' + data_string[1])
            data_list_dic[data_string[0]] = data_string[1]

            break

        else:

            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('flash_shading', data_list_dic)


def read_flash_noise(read_file):

    # Nosie
    # 270  Patch 	 VN (1)  	 VN (2) 	 VN (3)	 d_L (1)  	 d_L (2) 	 dL (3)	 d_u (1)  	 d_u (2) 	 d_u (3)
    # d_v (1)  	 d_v (2) 	 d_v (3)
    # 271     OECF-20 	 0.535 	 0.185 	 0.242	 0.483 	 0.170 	 0.222	 0.048 	 0.016 	 0.021	 0.033 	 0.005 	 0.007
    # 272     OECF-19 	 0.777 	 0.525 	 0.541	 0.626 	 0.394 	 0.405	 0.142 	 0.124 	 0.129	 0.094 	 0.077 	 0.082
    # 273     OECF-18 	 1.186 	 1.030 	 1.009	 1.071 	 0.932 	 0.910	 0.114 	 0.098 	 0.099	 0.055 	 0.043 	 0.045
    # 274     OECF-17 	 0.638 	 0.268 	 0.308	 0.592 	 0.248 	 0.284	 0.045 	 0.019 	 0.022	 0.026 	 0.013 	 0.014
    # 275     OECF-16 	 1.539 	 0.709 	 0.803	 1.404 	 0.643 	 0.726	 0.134 	 0.064 	 0.074	 0.065 	 0.037 	 0.041
    # 276     OECF-15 	 1.731 	 0.587 	 0.776	 1.594 	 0.541 	 0.715	 0.140 	 0.047 	 0.063	 0.057 	 0.017 	 0.023
    # 277     OECF-14 	 2.874 	 1.102 	 1.362	 2.651 	 1.018 	 1.258	 0.235 	 0.091 	 0.112	 0.070 	 0.020 	 0.025
    # 278     OECF-13 	 2.932 	 1.092 	 1.371	 2.599 	 0.862 	 1.119	 0.339 	 0.236 	 0.257	 0.138 	 0.091 	 0.104
    # 279     OECF-12 	 2.755 	 0.998 	 1.287	 2.408 	 0.757 	 1.020	 0.346 	 0.247 	 0.270	 0.160 	 0.095 	 0.111
    # 280     OECF-11 	 3.308 	 1.433 	 1.685	 3.066 	 1.325 	 1.558	 0.265 	 0.118 	 0.138	 0.050 	 0.024 	 0.030
    # 281     OECF-10 	 3.299 	 1.554 	 1.731	 3.029 	 1.429 	 1.589	 0.241 	 0.123 	 0.135	 0.198 	 0.066 	 0.084
    # 282     OECF-9 	 2.834 	 1.187 	 1.400	 2.586 	 1.077 	 1.269	 0.231 	 0.109 	 0.127	 0.160 	 0.052 	 0.072
    # 283     OECF-8 	 2.710 	 1.478 	 1.575	 2.443 	 1.319 	 1.402	 0.205 	 0.124 	 0.129	 0.284 	 0.164 	 0.197
    # 284     OECF-7 	 1.918 	 0.803 	 0.972	 1.614 	 0.582 	 0.726	 0.190 	 0.171 	 0.180	 0.439 	 0.232 	 0.288
    # 285     OECF-6 	 1.733 	 1.010 	 1.099	 1.293 	 0.632 	 0.698	 0.233 	 0.237 	 0.242	 0.748 	 0.541 	 0.605
    # 286     OECF-5 	 1.369 	 0.483 	 0.626	 1.274 	 0.449 	 0.582	 0.102 	 0.036 	 0.047	 0.027 	 0.009 	 0.012
    # 287     OECF-4 	 1.752 	 0.837 	 0.981	 1.298 	 0.476 	 0.591	 0.405 	 0.341 	 0.362	 0.338 	 0.219 	 0.252
    # 288     OECF-3 	 1.343 	 0.486 	 0.608	 1.234 	 0.446 	 0.558	 0.116 	 0.043 	 0.054	 0.031 	 0.011 	 0.015
    # 289     OECF-2 	 1.606 	 0.793 	 0.911	 1.246 	 0.523 	 0.611	 0.367 	 0.280 	 0.310	 0.147 	 0.097 	 0.112
    # 290     OECF-1 	 1.250 	 0.481 	 0.575	 1.128 	 0.428 	 0.510	 0.109 	 0.052 	 0.061	 0.091 	 0.030 	 0.042

    read_line = read_file.readline()
    count_line = 0
    n = 243 + 1  # 接着flash_shading后面继续读

    data_list = [[0 for col in range(4)] for row in range(20)]

    while read_line:

        if data_model.read_flash_noise_line_number[count_line] == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split('\t'):
                data_string.append(x)

            print(data_string[0] + ':' + data_string[3], data_string[6], data_string[9], data_string[12][:-1])

            data_list[count_line][0] = data_string[3][1:]
            data_list[count_line][1] = data_string[6][1:]
            data_list[count_line][2] = data_string[9][1:]
            data_list[count_line][3] = data_string[12][1:-1]

            count_line += 1
            n += 1

            if n > data_model.read_flash_noise_line_number[-1]:
                break

        else:
            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('flash_noise', cal_noise(data_list))


def read_color(read_file):

    read_line = read_file.readline()
    data_list_dic = {}
    count_line = 0
    n = 1 + 1

    while read_line:

        if data_model.read_color_line_number[count_line] == n:

            read_line = read_file.readline().strip()

            data_string = []

            for x in read_line.split(','):
                data_string.append(x)

            # print(data_string)
            if count_line < 3:

                print(data_string[0] + ' ' + data_model.color_data_type_dic[0] + ':' + data_string[9])
                data_list_dic[data_string[0] + data_model.color_data_type_dic[0]] = data_string[9][1:]

            elif 3 <= count_line < 27:

                print(data_model.color_data_type_dic[1] + ' ' + data_string[0] + ':' + data_string[1][1:])
                data_list_dic[data_model.color_data_type_dic[1] + data_string[0]] = data_string[1][2:]

            elif count_line >= 27:

                print(data_model.color_data_type_dic[2] + ': ' + data_string[1])
                data_list_dic[data_model.color_data_type_dic[2]] = data_string[1]

            count_line += 1
            n += 1

            if n > data_model.read_color_line_number[-1]:
                break

        else:

            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('color', data_list_dic)


def read_shading(read_file):

    read_line = read_file.readline()
    data_list_dic = {}
    count_line = 0
    n = 1 + 1

    while read_line:

        if data_model.read_shading_line_number[count_line] == n:

            read_line = read_file.readline()

            data_string = []

            for x in read_line.split(','):
                data_string.append(x)

            if count_line < 1:
                print(data_string[0] + ':' + data_string[1])
                data_list_dic[data_string[0]] = data_string[1][0:-1]

            elif 1 <= count_line < 3:
                print(data_model.shading_data_type[1] + ' ' + data_string[0] + ':' + data_string[2])
                print(data_model.shading_data_type[2] + ' ' + data_string[0] + ':' + data_string[-2])
                data_list_dic[data_model.shading_data_type[1] + data_string[0]] = data_string[2]
                data_list_dic[data_model.shading_data_type[2] + data_string[0]] = data_string[-2]

            count_line += 1
            n += 1

            if n > data_model.read_shading_line_number[-1]:
                break

        else:

            read_line = read_file.readline()
            n += 1

    data_model.set_data_dic('shading', data_list_dic)


def read_focus(focus_type, path):

    print(path)

    focus_data_dic = {}
    data_list = []
    x = 0
    while x < len(path):

        print(path[x])

        wb = txtOperate.excel.Workbooks.Open(path[x])
        ws = wb.Worksheets('sheet1')

        data_list.append(ws.Cells(85, 4).Value)
        txtOperate.excel.Application.Quit()

        x += 1

    focus_data_dic[focus_type] = data_list

    return focus_data_dic
