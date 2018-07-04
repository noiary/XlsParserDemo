# coding:utf-8
import xlrd
import configparser
import xlswrite

config = configparser.ConfigParser()
config.read('channels-config')

# config.add_section('Build-Config')
# print(config.get('time_default_sign_in','Build-Config'))
# time_sign_in =



xml_name = u'new_1_标准报表.xls'

workbook = xlrd.open_workbook('in/' + xml_name)
sheet_names = workbook.sheet_names()

dict = {}
# 姓名列表
list_name = []
# 日期字典，key对应某一天，value对应这一天的上下班打卡数据
dict_person = {}

# 遍历表
for sheet_name in sheet_names:
    list_name.clear()
    sheet = workbook.sheet_by_name(sheet_name)
    dict_person.clear()
    # 遍历行
    for row in range(sheet.nrows):
        values = sheet.row_values(row)
        # 遍历列
        for col in range(len(values)):
            cell = sheet.cell(row, col).value
            if cell == '部门':
                # 获取姓名
                list_name.append(sheet.cell(row, col + 9).value)
            elif (col == 0 or col == 15 or col == 30) and row > 10:
                # 获取日期单元格 如 [01 一]
                cell_date = sheet.cell(row, col).value
                list_date = []
                for i in range(13):
                    offset_col = i + 1
                    # 上班时间
                    if offset_col == 1 or offset_col == 10:
                        time_sign_in = sheet.cell(row, col + offset_col).value
                        if time_sign_in:
                            list_date.append(time_sign_in)
                            continue
                    # 下班时间
                    elif offset_col == 3 or offset_col == 12:
                        time_sign_out = sheet.cell(row, col + offset_col).value
                        if time_sign_out:
                            if len(list_date) == 0:
                                list_date.append('_')  # 上班漏打卡
                            list_date.append(time_sign_out)
                            break

                # 检查上下班时间是否有漏打卡
                if len(list_date) == 0:
                    list_date.append('_')
                if len(list_date) == 1 and not list_date[0] == '旷工':
                    list_date.append('_')  # 下班漏打卡
                # 每个表有三个人的打卡数据，每个人的日期作为dict的key不能重复，追加'_'避免重复
                while dict_person.get(cell_date):
                    cell_date += '_'
                #
                dict_person[cell_date] = list_date

    # 加入名字和考勤数据到字典
    len_list_name = len(list_name)
    for i in range(len_list_name):
        name = list_name[i]
        #  如果名字重复，名字后面追加一个"+"
        while name and dict.get(name):
            name += '+'
        list_date = list(dict_person.keys())
        len_date = len(list_date) // len_list_name
        # 月初到月末的每一天日期数据组成的列表 <class 'list'>: ['01 一', '02 二', '03 三' ...... '29 一', '30 二', '31 三']
        _list_date = [date for index, date in enumerate(list_date) if index % len_list_name == i]
        # 把名字和每一天的签到数据加入到总字典数据中 {'甄鹏': {'01 一': ['旷工'], ...},  ......  '苟铭钥': {... '31 三': ['08:48', '18:44']}}
        dict[name] = {sign.replace('_', ''): dict_person.get(sign) for sign in _list_date}

xlswrite.write('out/[new]' + xml_name, dict)
# for key in dict.keys():
#     value = dict.get(key)
#     print('=====================', key, '=====================', '\n')
#     for k in value.keys():
#         _l = list(value.get(k))
#         list_format = get_time(k, _l)
#         for s in list_format:
#             if len(s) != 0:
#                 print(s, end='\n')
