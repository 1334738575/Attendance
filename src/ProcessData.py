import xlrd  # 读Excel数据用
import tkinter as tk
from tkinter import filedialog
from xlutils.copy import copy

import WriteData

import os


init_path = os.getcwd()
setting_path = os.path.join(init_path, 'Setting')
result_path = os.path.join(init_path, 'Result')


# Folder_Path = filedialog.askdirectory() # 文件夹路径
def read_excel():
    # 处理路径
    # root = tk.Tk()
    # root.withdraw()
    File_Path = filedialog.askopenfilename()

    data = xlrd.open_workbook(File_Path)
    # data是Excel里的数据
    sheet = data.sheet_by_index(0)

    # 存储初始表格中的数据（名字，天数，日期）
    ex_datax = []
    for row in range(sheet.nrows):
        if row < 3:
            continue
        dates = sheet.cell_value(row, 3).split("、")
        # print(dates)
        num_date = 0
        i_date = 0
        for date in dates:
            if date == '':
                del dates[i_date]
                continue
            num_date += 1
            i_date += 1
        # print(num_date)
        if num_date == 0:
            continue
        # print(dates)
        ex_one = [sheet.cell_value(row, 1), num_date, dates]
        # print(ex_one)
        ex_datax.append(ex_one)
    # print(ex_datax)
    # print(ex_datax[0][0])
    return ex_datax


def get_infor():
    File_Path = os.path.join(setting_path, '系统录入模板.xls')
    # File_Path = 'G:/系统文件夹/桌面/work/学生处人员在翔及到翔安值班情况表/系统录入模板.xls'

    data = xlrd.open_workbook(File_Path)
    # data是Excel里的数据
    sheet = data.sheet_by_index(0)

    # 存储模板中的数据（名字，天数）
    ex_datax = []
    staff_nums = []
    now_staff_size = 0
    for row in range(sheet.nrows):
        if row < 1:
            continue
        name = sheet.cell_value(row, 1)
        staff_num = sheet.cell_value(row, 0)
        # print(staff_num)
        if staff_num == '':
            continue
        # temp_num = int(staff_num)
        temp_num = staff_num
        # print(ex_one)
        ex_datax.append(name)
        staff_nums.append(temp_num)
        now_staff_size += 1
    dic = dict(zip(ex_datax, staff_nums))

    return now_staff_size, dic


# def alignment_data(data_in, data_dic):
#     infor_data = []
#     new_infor = []
#     for one_in in data_in:
#         name = one_in[0]
#         infor_one = []
#         infor_one.append(name)
#         if name in data_dic:
#             staff_num = data_dic[name]
#         else:
#             print('未查询到*', name, '*的教工号!')
#             staff_num = input('请手动输入：')
#             new_one = []
#             new_one.append(name)
#             new_one.append(staff_num)
#             new_infor.append(new_one)
#         money = 25 * one_in[1]
#         infor_one.append(staff_num)
#         infor_one.append(money)
#         # print(infor_one)
#         infor_data.append(infor_one)
#     # print(infor_data)
#
#     return new_infor, infor_data


def update_infor(new_in, now_size):
    index = len(new_in)
    infor_path = os.path.join(setting_path, '系统录入模板.xls')
    infor_file = xlrd.open_workbook(infor_path)
    infor_sheet = infor_file['Sheet1']
    rows_old = infor_sheet.nrows
    new_book = copy(infor_file)
    new_sheet = new_book.get_sheet(0)
    new_sheet.col(0).width = 15 * 256
    new_sheet.col(1).width = 10 * 256
    for i in range(0, index):
        new_sheet.write(rows_old+i, 0, new_in[i][1])
        new_sheet.write(rows_old+i, 1, new_in[i][0])
    new_book.save(infor_path)


def ProcessData(the_year, the_month):
    ex_data = read_excel()
    WriteData.create_print(ex_data, the_year, the_month)
    now_staff_size, infor_dic = get_infor()
    return ex_data, infor_dic, now_staff_size


def ProcessData2(infor_data, new_infor, the_year, the_month, now_staff_size):
    WriteData.create_template(infor_data, the_year, the_month)
    if new_infor:
        update_infor(new_infor, now_staff_size)


if __name__ == '__main__':
    ex_data = read_excel()
    the_year = input('输入年份（如2022）：')
    the_month = input('输入月份（如1）：')
    # WriteData.create_print(ex_data, the_year, the_month)
    # now_staff_size, infor_dic = get_infor()
    # # ex_data.append(['lyj', 3, ''])
    # new_infor, infor_data = alignment_data(ex_data, infor_dic)
    # WriteData.create_template(infor_data, the_year, the_month)
    # if not new_infor:
    #     print('处理完毕，且无新人员更新！')
    # else:
    #     update_infor(new_infor, now_staff_size)

    # name1 = ex_data[0][0]
    # print(name1)
    # if name1 in infor_dic:
    #     print(infor_dic[name1])
    # else:
    #     print('未存储人员!')

    # copy_xlsx()
    # # debug
    # print_target_file = filedialog.askopenfilename()
    #
    # print_sheet_all = openpyxl.load_workbook(print_target_file)
    # print('open ', print_target_file)
    # print_sheet = print_sheet_all['Sheet1']
    # num = 1
    # for print_one in ex_data:
    #     print_sheet.cell(num + 2, 1, num)
    #     print_sheet.cell(num + 2, 2, str(print_one[0]))
    #     print_sheet.cell(num + 2, 3, str(print_one[1]))
    #     print_str = ''
    #     for print_date in print_one[2]:
    #         if print_str == '':
    #             print_str = print_str + str(print_date)
    #             continue
    #         print_str = print_str + '、' + str(print_date)
    #     print_sheet.cell(num + 2, 4, print_str)
    #     num += 1
    #     # print(print_one)
    # print_sheet_all.save(print_target_file)


