import xlwt
import time
import config

TAG = 'xmlwrite'


def write(out_name, my_dict):
    wb = xlwt.Workbook()
    ws = wb.add_sheet('A Test Sheet')
    row = 0;
    for name in my_dict.keys():
        value = my_dict.get(name)
        # print(TAG, 'name', name)
        if not name:
            continue
        # ws.write(row, 0, name)  # 名字
        for k in value.keys():
            _l = list(value.get(k))
            list_format = get_time(k, _l)
            for ymd_week_hm in list_format:
                ws.write(row, 0, name)
                for i, s in enumerate(ymd_week_hm):
                    print(TAG, 'name', name, 'row', row, ',\t', 'value', s)
                    ws.write(row, i + 1, s)
                row += 1
    wb.save(out_name)


# 接收'1 一' 返回'2017年6月1日 星期一'
def get_time(str_day_week, list_h_m):
    strftime = time.strftime("%Y/%m/", time.localtime())
    list_day_week = str_day_week.split(' ')
    week = list_day_week[1]

    list_result = []
    if week != '日' and week != '六':
        if list_h_m[0] == "旷工" and config.AUTO_SIGN_IF_NOT:
            list_h_m = ['09:30', '18:00']
        for index, hm in enumerate(list_h_m):
            list_y_m_d = [strftime + list_day_week[0], ' 星期' + week]
            if hm == '_' and config.AUTO_SIGN_IF_NOT:
                hm = '09:30' if index == 0 else '18:00'
            list_y_m_d.append(hm)
            list_result.append(list_y_m_d)
    else:
        for hm in list_h_m:
            if not hm == '_':
                list_y_m_d = [strftime + list_day_week[0], ' 星期' + week, hm]
                list_result.append(list_y_m_d)
    return list_result
