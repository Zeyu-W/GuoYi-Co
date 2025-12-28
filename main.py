# -*- coding: UTF-8 -*-
import tkinter
from tkinter.filedialog import askdirectory
from tkinter.messagebox import *
from tkinter import ttk
import os
import xlrd
import xlwt
from xlutils.copy import copy


def select_path_1():
    directory_1 = askdirectory()
    path_1.set(directory_1)


def select_path_2():
    directory_2 = askdirectory()
    path_2.set(directory_2)


class TableChecker:
    def __init__(self):
        self.source_dir = entry_1.get()
        self.target_dir = entry_2.get()
        self.source_files_num = len(os.listdir(self.source_dir))
        self.target_files_num = len(os.listdir(self.target_dir))

    def __len__(self):
        return self.source_files_num

    def __getitem__(self, item):
        source_file_name = os.listdir(self.source_dir)
        path = [os.path.join(self.source_dir, source_file_name[item]),
                os.path.join(self.target_dir, source_file_name[item])]
        path[0] = path[0].replace('/', '\\')
        path[1] = path[1].replace('/', '\\')
        return path


def main():
    checker = TableChecker()
    if checker.source_files_num == checker.target_files_num:
        current_value = 0
        progressbar["value"] = current_value
        max_value = len(checker)
        progressbar["maximum"] = max_value

        for i in range(len(checker)):
            path = checker[i]
            checking(path)
            progressbar["value"] = i + 1
            progressbar.update()
    else:
        showerror("警告", "文件不匹配")


def checking(path):
    source = xlrd.open_workbook(path[0])
    target = xlrd.open_workbook(path[1], formatting_info=True)
    names_source = source.sheet_by_index(0).col_values(1)
    title_source = title_creator(source.sheet_by_index(0))
    names_target = target.sheet_by_index(0).col_values(1, 4, 90)
    title_target = title_creator(target.sheet_by_index(0))[3:]

    write_buffer = copy(target)
    write_sheet_buffer = write_buffer.get_sheet(0)

    for row_position_in_target in range(len(names_target)):
        try:
            row_position_in_source = names_source.index(names_target[row_position_in_target])
        except ValueError:
            pass
        for column_position_in_target in range(len(title_target)):
            try:
                column_position_in_source = title_source.index(title_target[column_position_in_target])
                value = source.sheet_by_index(0).cell_value(row_position_in_source, column_position_in_source)
                write_sheet_buffer.write(row_position_in_target + 4, column_position_in_target + 3, value)
            except ValueError:
                pass
    write_buffer.save('test.xls')



def title_creator(sheet):
    title_0 = sheet.row_values(2)
    title_1 = sheet.row_values(3)
    title = [None] * len(title_0)
    for i in range(len(title_0)):
        if title_0[i] == '':
            title[i - 1] = title[i - 1][0:2] + title_1[i - 1]
            title[i] = title[i - 1][0:2] + title_1[i]
        else:
            title[i] = title_0[i]
    return title

root = tkinter.Tk()
root.title('guoyi Co')
root.geometry("515x150+500+300")
root.resizable(False, False)
path_1 = tkinter.StringVar()
path_2 = tkinter.StringVar()

label_1 = tkinter.Label(root, text="选择工资表路径:", font=("Arial", 12), width=18, height=2)
label_2 = tkinter.Label(root, text="选择模板路径:", font=("Arial", 12), width=18, height=2)
entry_1 = tkinter.Entry(root, textvariable=path_1, width=30)
entry_2 = tkinter.Entry(root, textvariable=path_2, width=30)
button_1 = tkinter.Button(root, text="选择", command=select_path_1, font=("Arial", 12), width=10, height=1)
button_2 = tkinter.Button(root, text="选择", command=select_path_2, font=("Arial", 12), width=10, height=1)
button_3 = tkinter.Button(root, text="开始核对", command=main, font=("Arial", 12), width=10, height=1)
progressbar = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")

label_1.grid(row=0, column=0, sticky=tkinter.W)
entry_1.grid(row=0, column=1)
button_1.grid(row=0, column=2, padx=10)
label_2.grid(row=1, column=0, sticky=tkinter.W)
entry_2.grid(row=1, column=1)
button_2.grid(row=1, column=2, padx=10)
button_3.grid(row=2, column=2, pady=10)
progressbar.grid(row=2, column=0, columnspan=2, sticky=tkinter.E)
root.mainloop()
