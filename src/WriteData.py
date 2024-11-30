import xlwt
import tkinter as tk
from tkinter import filedialog

# import win32com.client as win32
import os


init_path = os.getcwd()
setting_path = os.path.join(init_path, 'Setting')
result_path = os.path.join(init_path, 'Result')
# root = tk.Tk()
# root.withdraw()


# def xls_to_xlsx(xls_name):
#     file_name = xls_name
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(file_name)
#     wb.SaveAs(file_name+"x", FileFormat=51)    #FileFormat = 51 is for .xlsx extension
#     wb.Close()                               #FileFormat = 56 is for .xls extension
#     excel.Application.Quit()


def create_xls():
    workbook = xlwt.Workbook(encoding='utf-8')
    sheet = workbook.add_sheet('sheet1')
    return workbook, sheet


def write_to_sheet(xls_name, data_in, the_year, the_month):
    workbook, sheet = create_xls()
    all_num = 0
    for print_one in data_in:
        all_num += 1
        # print(print_one)
    max_row = all_num + 2
    # 行高列宽----------------------------------------
    # 格式：sheet.col(n).width = 11 * 256 ，表示第n列的宽度为11个字符
    sheet.col(0).width = 12 * 256  # 实际转换关系还不清楚
    sheet.col(1).width = 12 * 256
    sheet.col(2).width = 10 * 256
    sheet.col(3).width = 10 * 256
    sheet.col(4).width = 12 * 256
    sheet.col(5).width = 9 * 256
    sheet.col(6).width = 13 * 256
    sheet.col(7).width = 14 * 256
    # 设置第0行的高度为800
    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 660  # 实际值为height/20（磅）
    sheet.row(1).height_mismatch = True
    sheet.row(1).height = 500
    for i in range(2, max_row):
        sheet.row(i).height_mismatch = True
        sheet.row(i).height = 540
    sheet.row(max_row).height_mismatch = True
    sheet.row(max_row).height = 720

    # 单元格基础格式（可复用）
    # 设置字体格式----------------------------
    font1 = xlwt.Font()
    font1.name = '黑体'
    font1.height = 20 * 20  # 一个20为单位，另一个为实际字号
    font1.bold = False
    # 设置边框格式--------------------------
    borders1 = xlwt.Borders()
    # 细实线:1，小粗实线:2，细虚线:3，中细虚线:4，大粗实线:5，双线:6，细点虚线:7
    # 大粗虚线:8，细点划线:9，粗点划线:10，细双点划线:11，粗双点划线:12，斜点划线:13
    borders1.left = 1
    borders1.right = 1
    borders1.top = 1
    borders1.bottom = 1
    # 设置对齐方式------------------------------
    alignment1 = xlwt.Alignment()
    # 0x01(左端对齐)、0x02(水平方向上居中对齐)、0x03(右端对齐)
    alignment1.horz = 0x02
    # 0x00(上端对齐)、 0x01(垂直方向上居中对齐)、0x02(底端对齐)
    alignment1.vert = 0x01

    # 标题--------------------------------------------
    style0 = xlwt.XFStyle()
    style0.font = font1
    style0.borders = borders1
    style0.alignment = alignment1
    # 合并单元格（r1，r2，c1，c2，v）
    sheet.write_merge(0, 0, 0, 7, '学生处人员到翔安校区值班情况表（' + the_year + '年' + the_month + '月份）', style=style0)
    # 落款--------------------------------------------
    font2 = xlwt.Font()
    font2.name = '宋体'
    font2.height = 20*11
    font2.bold = False
    style1 = xlwt.XFStyle()
    style1.borders = borders1
    style1.alignment = alignment1
    style1.font = font2
    sheet.write(max_row, 0, '经办人签名', style=style1)
    sheet.write_merge(max_row, max_row, 1, 3, style=style1)
    sheet.write(max_row, 4, '审核人签名', style=style1)
    sheet.write_merge(max_row, max_row, 5, 7, style=style1)
    # 表头--------------------------------------------
    font3 = xlwt.Font()
    font3.name = '宋体'
    font3.height = 20 * 11  # 一个20为单位，另一个为实际字号
    font3.bold = False
    style2 = xlwt.XFStyle()
    style2.font = font3
    style2.borders = borders1
    alignment2 = xlwt.Alignment()
    alignment2.horz = 0x02
    alignment2.vert = 0x01
    style2.alignment = alignment2
    style2.alignment.wrap = 1
    sheet.write(1, 0, '序号', style=style2)
    sheet.write(1, 1, '人员姓名', style=style2)
    sheet.write(1, 2, '值班天数', style=style2)
    sheet.write_merge(1, 1, 3, 7, '值班日期', style=style2)
    # 内容--------------------------------------------
    for i in range(2, max_row):
        print_str = ''
        for print_data in data_in[i-2][2]:
            if print_str == '':
                print_str = print_str + print_data
                continue
            print_str = print_str + '、' + print_data
        sheet.write_merge(i, i, 3, 7, print_str, style=style2)
        sheet.write(i, 0, i-1, style=style2)
        sheet.write(i, 1, data_in[i-2][0], style=style2)
        sheet.write(i, 2, data_in[i-2][1], style=style2)

    workbook.save(xls_name)


def create_print(data_in, the_year, the_month):
    # Folder_Path = filedialog.askdirectory()  # 文件夹路径
    # Folder_Path = 'G:/系统文件夹/桌面/work/学生处人员到翔安值班情况表/'+str(the_year)+'.'+str(the_month)+'-'+str(the_month+1)
    # isExists = os.path.exists('G:/系统文件夹/桌面/work/学生处人员到翔安值班情况表/'+str(the_year)+'.'+str(the_month)+'-'+str(the_month+1))
    Folder_Path = os.path.join(result_path, the_year + '.' + the_month)
    isExists = os.path.exists(Folder_Path)
    if not isExists:
        # os.path.exists(path+str(i)) 创建文件夹 路径+名称
        os.makedirs(Folder_Path)
    xls_name = Folder_Path + '/' + str(the_year) + '.' + str(the_month) + '.xls'
    write_to_sheet(xls_name, data_in, the_year, the_month)


def write_to_infor(xls_name, infor_data):
    workbook, sheet = create_xls()
    all_num = 0
    for print_one in infor_data:
        all_num += 1
        # print(print_one)
    max_row = all_num + 1
    # # 行高列宽----------------------------------------
    # # 格式：sheet.col(n).width = 11 * 256 ，表示第n列的宽度为11个字符
    sheet.col(0).width = 15 * 256  # 实际转换关系还不清楚
    sheet.col(1).width = 10 * 256
    # sheet.col(2).width = 10 * 256
    # sheet.col(3).width = 10 * 256
    # sheet.col(4).width = 12 * 256
    # sheet.col(5).width = 9 * 256
    # sheet.col(6).width = 13 * 256
    # sheet.col(7).width = 14 * 256
    # # 设置第0行的高度为800
    # sheet.row(0).height_mismatch = True
    # sheet.row(0).height = 660  # 实际值为height/20（磅）
    # sheet.row(1).height_mismatch = True
    # sheet.row(1).height = 500
    # for i in range(2, max_row):
    #     sheet.row(i).height_mismatch = True
    #     sheet.row(i).height = 540
    # sheet.row(max_row).height_mismatch = True
    # sheet.row(max_row).height = 720

    # 单元格基础格式（可复用）
    # 设置字体格式----------------------------

    # 表头--------------------------------------------
    font3 = xlwt.Font()
    font3.name = '宋体'
    font3.height = 20 * 12  # 一个20为单位，另一个为实际字号
    font3.bold = False
    style2 = xlwt.XFStyle()
    style2.font = font3
    sheet.write(0, 0, '工号', style=style2)
    sheet.write(0, 1, '人员姓名', style=style2)
    sheet.write(0, 2, '其他', style=style2)
    # style3 = xlwt.XFStyle()
    # style3.font = font3

    # 内容--------------------------------------------
    for i in range(1, max_row):
        sheet.write(i, 0, infor_data[i-1][1], style=style2)
        # print(infor_data[i-1][1])
        sheet.write(i, 1, infor_data[i-1][0], style=style2)
        sheet.write(i, 2, infor_data[i-1][2], style=style2)

    workbook.save(xls_name)


def create_template(infor_data, the_year, the_month):
    folder_path = os.path.join(result_path, the_year + '.' + the_month)
    isExists = os.path.exists(folder_path)
    if not isExists:
        # os.path.exists(path+str(i)) 创建文件夹 路径+名称
        os.makedirs(folder_path)
    xls_name = folder_path + '/' + str(the_year) + '.' + str(the_month) + '系统录入.xls'
    write_to_infor(xls_name, infor_data)


if __name__ == '__main__':
    data = []
    create_print(data, 2022, 6)
    # Folder_Path = filedialog.askdirectory()  # 文件夹路径
    # xls_name = Folder_Path + '/' + str(3) + '.' + str(5) + '.xls'
    # write_to_sheet()
    # xls_to_xlsx()
    print('Finish!')
