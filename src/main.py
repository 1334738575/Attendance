# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.
from tkinter import *
import os
from tkinter.simpledialog import askinteger

from ProcessData import ProcessData, ProcessData2

init_path = os.getcwd()

# bool_click = False


# def Click():
#     global bool_click
#     bool_click = True


# class Page_1:  # 这是第一个页面
#     def __init__(self, window):
#         self.window = window
#         self.window.title("p1")
#         self.window.geometry("200x200")
#         self.window.config(bg="#F9C03D")
#         self.button = Button(self.window, text="确定", command=self.change)
#         self.button.pack()
#
#     def change(self):
#         # pass  # 不知道怎么写，先占位
#         self.button.destroy()


class AS_GUI:

    def __init__(self):
        self.as_win = Tk()
        self.as_win.title('此应用根据已统计的考勤表生成下一步文件')
        self.as_win.geometry('480x280')

        # 输入
        label_year = Label(self.as_win, text='输入年份', font=('宋体', 10), width=30, height=1)
        label_year.grid(row=0, column=1, padx=5, pady=5)
        label_month = Label(self.as_win, text='输入月份', font=('宋体', 10), width=30, height=1)
        label_month.grid(row=0, column=2, padx=5, pady=5)
        entry_year = Entry(self.as_win)
        entry_year.grid(row=1, column=1, padx=5, pady=5)
        entry_year.insert(0, '2023')
        self.entry_year = entry_year
        entry_month = Entry(self.as_win)
        entry_month.grid(row=1, column=2, padx=5, pady=5)
        entry_month.insert(0, '2')
        self.entry_month = entry_month

        # 选择已统计的考勤表
        button_choose = Button(self.as_win, text="选择已统计的考勤表", command=self.Start_Process, width=25, height=3)
        button_choose.grid(row=2, column=1, columnspan=3, padx=10, pady=5)
        
        label_number = Label(self.as_win, text='请选择文件，选择后即处理！', font=('宋体', 12), width=40, height=2)
        label_number.grid(row=5, column=1, columnspan=3, padx=5, pady=5)

        # 输入未知的教工号
        # label_name = Label(self.as_win, text='教工号未知人员：', font=('宋体', 10), width=30, height=1)
        # label_name.grid(row=3, column=1, padx=5, pady=5)
        # label_number = Label(self.as_win, text='请输入教工号', font=('宋体', 10), width=30, height=1)
        # label_number.grid(row=4, column=1, padx=5, pady=5)
        # # entry_name = Entry(self.as_win)
        # # entry_name.grid(row=3, column=2, padx=5, pady=5)
        # # entry_name.insert(0, '2023')
        # entry_number = Entry(self.as_win)
        # entry_number.grid(row=4, column=2, padx=5, pady=5)
        # self.entry_number = entry_number
        # entry_number.insert(0, '2')
        # button_choose = Button(self.as_win, text="确定", command=Click, width=10)
        # button_choose.grid(row=5, column=2, padx=5, pady=5)

        # 打开结果
        button_choose = Button(self.as_win, text="打开结果", command=OpenResult, width=15, height=2)
        button_choose.grid(row=7, column=1, columnspan=3, padx=5, pady=5)

        # 关闭
        button_close = Button(self.as_win, text="关闭", command=self.as_win.destroy, width=10)
        button_close.grid(row=8, column=0, columnspan=3, padx=5, pady=5)

        self.as_win.mainloop()

    def Start_Process(self):
        the_year = self.entry_year.get()
        the_month = self.entry_month.get()
        ex_data, infor_dic, now_staff_size = ProcessData(the_year, the_month)
        if now_staff_size==0:
            label_number = Label(self.as_win, text='文件处理失败，请确认数据格式、文件后缀xls、配置文件路径!!!', font=('宋体', 12), width=50, height=1)
            label_number.grid(row=5, column=1, columnspan=3, padx=5, pady=5)
            return
        # self.ex_data = ex_data
        # self.infor_dic = infor_dic
        new_infor, infor_data = self.alignment_data(ex_data, infor_dic)
        ProcessData2(infor_data, new_infor, the_year, the_month, now_staff_size)
        label_number = Label(self.as_win, text='文件处理完毕!!!', font=('宋体', 12), width=50, height=1)
        label_number.grid(row=5, column=1, columnspan=3, padx=5, pady=5)

    def alignment_data(self, data_in, data_dic):
        global bool_click
        infor_data = []
        new_infor = []
        for one_in in data_in:
            name = one_in[0]
            infor_one = []
            infor_one.append(name)
            if name in data_dic:
                staff_num = data_dic[name]
            else:
                # label_number = Label(self.as_win, text=name, font=('宋体', 10), width=30, height=1)
                # label_number.grid(row=4, column=2, padx=5, pady=5)
                # while not bool_click:
                #     i = 0
                # bool_click = False
                staff_num = askinteger(title='请输入该人员的教工号',
                                       prompt=name+':')
                # staff_num = self.entry_number.get()
                new_one = []
                new_one.append(name)
                new_one.append(staff_num)
                new_infor.append(new_one)
            money = 25 * one_in[1]
            infor_one.append(staff_num)
            infor_one.append(money)
            # print(infor_one)
            infor_data.append(infor_one)
        # print(infor_data)

        return new_infor, infor_data


def OpenResult():
    path = os.path.join(init_path, 'Result')
    os.startfile(path)


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    AS_GUI()

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
