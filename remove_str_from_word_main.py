# -*- coding: utf-8 -*-
###########################################################################
#    Copyright 2023 奇客罗方智能科技 https://www.geekercloud.com
#    ITGeeker.net <alanljj@gmail.com>
############################################################################
import os
import shutil
import tkinter
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import tkinter.filedialog
from remove_str_api import generate_file_and_str_list

class ListSheet(tk.Frame):

    def __init__(self, master):
        super().__init__(master)

        self.label1 = tk.Label(self, text='待移除文字列表')
        self.label1.config(font=('Microsoft YaHei UI', 14))
        self.label1.grid(row=0, column=0, columnspan=2, ipadx=10, ipady=10)

        # , selectmode = MULTIPLE
        self.listbox = tk.Listbox(self, width=66, font=('Microsoft YaHei UI', 12))
        # self.listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.listbox.grid(row=1, column=0, columnspan=2, padx=10, pady=10)

        # self.entry1 = tk.Entry(self)
        # input_text = tk.StringVar(self, value='输入您想移除的字符')
        # self.entry1 = ttk.Entry(self, textvariable=input_text, justify=CENTER)
        self.entry1 = tk.Entry(self, justify=CENTER, width=60,
                               font=('Microsoft YaHei UI', 12))
        # self.init_placeholder(self.entry1, "Enter your text here")
        # self.entry1.insert(0, "输入您想移除的字符")
        self.entry1.focus_force()
        # self.entry1.pack(padx=10, pady=10, side=TOP, ipadx=30, ipady=6)
        self.entry1.grid(row=2, column=0, columnspan=2, padx=10, pady=10)
        # self.entry1.pack()
        # self.canvas1.create_window(200, 140, window=self.entry1)

        # self.add_button = tk.Button(self, text="添加要移除的字符", command=self.add_item, padx='5', pady='5')
        self.add_button = tk.Button(self, text="添加字符", command=self.add_item, bg='brown', fg='white',
                                    font=('Microsoft YaHei UI', 11, 'bold'))
        # self.add_button.pack(padx='5', pady='3', side="left", fill=tk.BOTH, expand=True)
        self.add_button.grid(row=3, column=0, padx=10, pady=10, ipadx=10, ipady=5)
        # self.add_button.grid(row=0, column=0)

        self.edit_button = tk.Button(self, text="编辑或删除", command=self.edit_item, bg='grey', fg='white',
                                     font=('Microsoft YaHei UI', 11, 'normal'))
        # self.edit_button.pack(padx='5', pady='3', side="left", fill=tk.BOTH, expand=True)
        self.edit_button.grid(row=3, column=1, padx=10, pady=10, ipadx=10, ipady=5)
        # self.edit_button.grid(row=0, column=1)

        # path of docx
        self.entry_path = ttk.Entry(self, justify=LEFT, width=60,
                                    font=('Microsoft YaHei UI', 11))
        self.entry_path.insert(0, "浏览并选择文件目录")
        # self.entry_path.focus_force()
        # self.entry_path.pack(padx=10, pady=10, side=TOP, ipadx=30, ipady=6)
        self.entry_path.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

        self.browse_button = tk.Button(self, text="选择目录", command=self.select_directory, bg='grey', fg='white',
                                       font=('Microsoft YaHei UI', 11, 'normal'))
        # self.browse_button.pack(padx='5', pady='3')
        self.browse_button.grid(row=5, column=0, padx=10, pady=10, ipadx=10, ipady=5)

        self.start_remove_button = tk.Button(self, text="开始处理", command=self.start_remove_strings_from_files,
                                             bg='purple',
                                             fg='white',
                                             font=('Microsoft YaHei UI', 11, 'bold'))
        # self.start_remove_button.pack(padx='5', pady='3')
        self.start_remove_button.grid(row=5, column=1, padx=10, pady=10, ipadx=10, ipady=5)

        self.read_all_item_to_list_box()

        win.protocol("WM_DELETE_WINDOW", self.on_window_close)

    # def init_placeholder(self, widget, placeholder_text):
    #     widget.placeholder = placeholder_text
    #     if widget.get() == "":
    #         widget.insert("end", placeholder_text)
    #
    #     def remove_placeholder(event):
    #         placeholder_text = getattr(event.widget, "placeholder", "")
    #         if placeholder_text and event.widget.get() == placeholder_text:
    #             event.widget.delete(0, "end")
    #
    #     widget.bind("<FocusIn>", remove_placeholder)
    #     widget.bind("<FocusOut>", self.init_placeholder)

    def add_item(self):
        # item = input("Enter item: ")
        item = self.entry1.get()
        self.listbox.insert(tk.END, item.strip())
        self.entry1.delete(0, END)

    def edit_item(self):
        if self.listbox.curselection():
            for item in self.listbox.curselection():
                print('self.listbox.get(item): ', self.listbox.get(item))
                self.entry1.insert(END, self.listbox.get(item))
                self.listbox.delete(item)
                # self.listbox.insert("end", "foo")
        else:
            messagebox.showwarning(title="Error Reminder", message="请先选择想要修改的文字！")

    def get_all_item_list(self):
        values = self.listbox.get(0, END)
        print('values: %s' % type(values))  # <class 'tuple'>
        print('values: %s' % list(values))
        return list(values)

    def start_remove_strings_from_files(self):
        val_list = self.get_all_item_list()
        if val_list:
            self.save_all_item_to_txt(val_list)
        if not val_list:
            messagebox.showwarning(title="Error Reminder", message="请添加要移除的文字！")
        elif not self.entry_path.get() or self.entry_path.get() == '浏览并选择文件目录':
            messagebox.showwarning(title="Error Reminder", message="请先选择文件的目录！")
        else:
            print('self.entry_path.get(): %s' % self.entry_path.get())
            finished = generate_file_and_str_list(self.entry_path.get(), val_list)
            if finished:
                messagebox.showinfo(title="任务通知", message="任务已圆满完成！")
            else:
                messagebox.showerror(title="任务错误通知", message="任务完成，但有错误！")

    def generate_string_text_ffp(self):
        cur_usr_path = os.environ['USERPROFILE']
        print('cur_usr_path: %s' % cur_usr_path)
        remove_str_f = os.path.join(cur_usr_path, 'remove_string.txt')
        if not os.path.isfile(remove_str_f):
            if os.path.isfile('staff_mobile_email.txt'):
                shutil.move('staff_mobile_email.txt', remove_str_f)
            else:
                with open(remove_str_f, 'w') as fp:
                    pass
        return remove_str_f

    def save_all_item_to_txt(self, value_list):
        remove_str_f = self.generate_string_text_ffp()
        with open(remove_str_f, 'w', encoding='utf-8') as f:
            for val in value_list:
                f.write(f"{val}\n")

    def read_all_item_to_list_box(self):
        remove_str_f = self.generate_string_text_ffp()
        with open(remove_str_f, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            print('lines: %s' % lines)
        if lines:
            for i in range(len(lines)):
                lines[i] = lines[i].rstrip("\n")
            for line in lines:
                if line:
                    self.listbox.insert(END, line)

    # def select_file(self):
    #     """Opens a file dialog and sets the entry widget to the selected path."""
    #     path = tk.filedialog.askopenfilename(
    #         parent=win,
    #         title="Choose a file",
    #         initialdir="D:\\",
    #     )
    #     self.entry_path.delete(0, tk.END)
    #     self.entry_path.insert(0, path)

    def select_directory(self):
        directory = tk.filedialog.askdirectory()
        self.entry_path.delete(0, END)
        self.entry_path.insert(0, directory)

    def on_window_close(self):
        print("Window closed")
        val_list = self.get_all_item_list()
        if val_list:
            self.save_all_item_to_txt(val_list)
        win.destroy()


if __name__ == "__main__":
    win = tk.Tk()
    win.eval('tk::PlaceWindow . center')
    win.iconbitmap(r"geekercloud_orange32.ico")

    win.title("技术奇客小工具-文档字符移除")
    # win.geometry("500x580")
    # ListSheet(win)
    list_sheet = ListSheet(win)
    list_sheet.pack()
    win.mainloop()
