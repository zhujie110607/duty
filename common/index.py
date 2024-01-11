# -*- 公共方法 -*-
import tkinter as tk
from tkinter import filedialog,messagebox as msgbox


"""
   检查文件夹是否存在，如果不存在则创建。
   :param folder_path_list: 文件夹路径
   """
import os


def create_folder_if_not_exists(folder_path_lest):
    for folder_path in folder_path_lest:
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

"""
    打开文件资源管理器，选择文件后，返回文件路径
    :param  prompt_message: 提示信息
"""


def select_excel_file(prompt_message):
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(filetypes=[(prompt_message, '*.xlsx')])

    if file_path:
        return os.path.abspath(file_path)
    else:
        show_message('没有选择文件', 0)
        return None

def show_message(message, x):
    if x == 0:
        msgbox.showerror('错误提示', message)
    else:
        msgbox.showinfo('温馨提示', message)