# -*- coding: utf-8 -*-
###########################################################################
#    Copyright 2023 奇客罗方智能科技 https://www.geekercloud.com
#    ITGeeker.net <alanljj@gmail.com>
############################################################################
import base64
import json
import os
import shutil
import tempfile
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import tkinter.filedialog
from webbrowser import open_new_tab

from remove_str_api import generate_file_and_str_list


class ListSheet(tk.Frame):

    def __init__(self, master):
        super().__init__(master)
        # self.label1 = tk.Label(self, text='待移除文字列表')
        # self.label1.config(font=('Microsoft YaHei UI', 14))
        # self.label1.grid(row=0, column=0, columnspan=2, ipadx=10, ipady=10)

        self.label_doc_nmb = None
        self.label_docx_nmb = None
        self.folder_frame = None
        self.start_remove_button = None
        self.browse_button = None
        self.edit_button = None
        self.listbox = None
        self.entry_path = None
        self.add_button = None
        self.entry_str = None

        self.list_frame()
        self.string_frame()
        self.select_path_frame()
        self.author_frame()

    def add_item(self):
        # item = input("Enter item: ")
        item = self.entry_str.get()
        self.listbox.insert(tk.END, item.strip())
        self.entry_str.delete(0, tk.END)

    def edit_item(self):
        if self.listbox.curselection():
            for item in self.listbox.curselection():
                print('self.listbox.get(item): ', self.listbox.get(item))
                self.entry_str.insert(tk.END, self.listbox.get(item))
                self.listbox.delete(item)
                # self.listbox.insert("end", "foo")
        else:
            messagebox.showwarning(title="Error Reminder", message="请先选择想要修改的文字！")

    def get_all_item_list(self):
        values = self.listbox.get(0, tk.END)
        print('values: %s' % type(values))  # <class 'tuple'>
        print('values: %s' % list(values))
        return list(values)

    def start_remove_strings_from_files(self):
        val_list = self.get_all_item_list()
        if val_list:
            self.save_all_item_to_json(val_list)
        if not val_list:
            messagebox.showwarning(title="Error Reminder", message="请添加要移除的文字！")
        elif not self.entry_path.get() or self.entry_path.get() == '浏览并选择目录':
            messagebox.showwarning(title="Error Reminder", message="请先选择文件的目录！")
        else:
            print('self.entry_path.get(): %s' % self.entry_path.get())
            finished_nmb = generate_file_and_str_list(self.entry_path.get(), val_list)
            if finished_nmb:
                messagebox.showinfo(title="任务通知", message="任务已圆满完成！共处理了%s个文件,\n"
                                                              "已处理文件带有-revised字样，保存在子目录《已处理文件》中。"
                                                              % str(finished_nmb))
            else:
                messagebox.showerror(title="任务错误通知", message="任务完成，但有错误！")

    def generate_json_ffp(self):
        cur_usr_path = os.environ['USERPROFILE']
        print('cur_usr_path: %s' % cur_usr_path)
        remove_str_f_old = os.path.join(cur_usr_path, 'staff_mobile_email.txt')
        remove_str_f = os.path.join(cur_usr_path, 'itgeeker_remove_str_from_word.json')
        if not os.path.isfile(remove_str_f):
            ffp_d = dict()
            if os.path.isfile('staff_mobile_email.txt'):
                old_str_l = []
                with open('staff_mobile_email.txt', 'r') as f:
                    lines = f.readlines()
                    if lines:
                        for i in range(len(lines)):
                            lines[i] = lines[i].rstrip("\n")
                        for line in lines:
                            if line:
                                old_str_l.append(line)
                if old_str_l:
                    ffp_d['string_list'] = old_str_l
                    with open(remove_str_f, 'w', encoding='utf-8') as jf:
                        jf.write(json.dumps(ffp_d, indent=4, ensure_ascii=False))
                shutil.move('staff_mobile_email.txt', remove_str_f_old)
            else:
                with open(remove_str_f, 'w', encoding='utf-8') as fp:
                    fp.write(json.dumps(ffp_d, indent=4, ensure_ascii=False))
                    # pass
        return remove_str_f

    def save_all_item_to_json(self, value_list):
        print("here should to save all")
        ffp_d = dict()
        remove_str_f = self.generate_json_ffp()

        if self.entry_path.get():
            ffp_d['entry_path'] = self.entry_path.get()

        print('ffp_d: ', ffp_d)
        with open(remove_str_f, 'w', encoding='utf-8') as ffp:
            ffp_d['string_list'] = value_list
            ffp.write(json.dumps(ffp_d, indent=4, ensure_ascii=False))

    def read_all_item_to_list_box(self):
        remove_str_f = self.generate_json_ffp()
        with open(remove_str_f, 'r', encoding='utf-8') as ffp:
            dt_dict = json.load(ffp)
            if 'string_list' in dt_dict:
                lines = dt_dict['string_list']
                print('lines: %s' % lines)
                if lines:
                    for i in range(len(lines)):
                        lines[i] = lines[i].rstrip("\n")
                    for line in lines:
                        if line:
                            self.listbox.insert(tk.END, line)
            if 'entry_path' in dt_dict:
                self.entry_path.delete(0, tk.END)
                self.entry_path.insert(0, dt_dict['entry_path'])
                self.folder_info_fram()
                self.cout_nmb_of_doc(dt_dict['entry_path'])

    def cout_nmb_of_doc(self, directory):
        count_docx = 0
        count_doc = 0
        for root_dir, cur_dir, files in os.walk(directory):
            for file in files:
                if file.endswith(".docx"):
                    count_docx += 1
                elif file.endswith(".doc"):
                    count_doc += 1
        print('file count_docx:', count_docx)
        print('file count_doc:', count_doc)
        self.label_docx_nmb.config(text='.docx文件：' + str(count_docx) + '个')
        self.label_doc_nmb.config(text='.doc文件：' + str(count_doc) + '个 (本工具只支持处理docx格式文件)')
        if count_doc:
            label_doc_reminder = ttk.Label(self.folder_frame,
                                           text='建议使用技术奇客的开源工具 - Word格式转换(.doc ➡ .docx) [点击下载并安装]',
                                           cursor="hand2")
            label_doc_reminder.config(font=('Microsoft YaHei UI', 10))
            label_doc_reminder.bind("<Button-1>",
                                    lambda e: self.open_website(
                                        "https://www.itgeeker.net/itgeeker-technical-service"
                                        "/itgeeker_convert_doc_to_docx/"))
            label_doc_reminder.grid(row=1, column=0, columnspan=2, padx=(15, 10), ipadx=10, ipady=5, sticky="ew")
    def select_directory(self):
        directory = tk.filedialog.askdirectory()
        self.entry_path.delete(0, tk.END)
        self.entry_path.insert(0, directory)
        self.folder_info_fram()
        self.cout_nmb_of_doc(directory)

    def open_website(self, url):
        open_new_tab(url)

    def on_window_close(self):
        print("Window closed")
        val_list = self.get_all_item_list()
        # if val_list:
        self.save_all_item_to_json(val_list)
        geekerWin.destroy()

    def list_frame(self):
        string_frame = ttk.LabelFrame(self, text="待移除文字列表")
        string_frame.grid(row=0, column=0, columnspan=2, sticky='nsew')
        # , selectmode = MULTIPLE
        self.listbox = tk.Listbox(string_frame, width=66, font=('Microsoft YaHei UI', 12))
        # self.listbox.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)
        self.listbox.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

        tree_y_scroll = ttk.Scrollbar(self.listbox, orient='vertical', command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=tree_y_scroll.set)
        tree_y_scroll.place(relx=1, rely=0, relheight=1, anchor='ne')
        # mousewheel scrolling
        self.listbox.bind('<MouseWheel>', lambda event: self.listbox.yview_scroll(-int(event.delta / 60), "units"))

    def string_frame(self):
        string_frame = ttk.LabelFrame(self, text="文字调整")
        string_frame.grid(row=2, column=0, columnspan=2, padx=10, pady=10, ipadx=10, sticky='nsew')

        self.entry_str = tk.Entry(string_frame, justify=tk.CENTER, width=62,
                                  font=('Microsoft YaHei UI', 12))
        # self.init_placeholder(self.entry_str, "Enter your text here")
        # self.entry_str.insert(0, "输入您想移除的字符")
        self.entry_str.focus_force()
        # self.entry_str.pack(padx=10, pady=10, side=TOP, ipadx=30, ipady=6)
        self.entry_str.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='ew')

        self.add_button = tk.Button(string_frame, text="添加字符", command=self.add_item, bg='brown', fg='white',
                                    font=('Microsoft YaHei UI', 11, 'bold'))
        self.add_button.grid(row=1, column=0, padx=10, pady=5, ipadx=10, ipady=5)

        self.edit_button = tk.Button(string_frame, text="编辑或删除", command=self.edit_item, bg='grey', fg='white',
                                     font=('Microsoft YaHei UI', 11, 'normal'))
        self.edit_button.grid(row=1, column=1, padx=10, pady=5, ipadx=10, ipady=5)

    def folder_info_fram(self):
        self.folder_frame = ttk.LabelFrame(self, text="目录信息")
        self.folder_frame.grid(row=4, column=0, columnspan=2, padx=10, pady=10, ipadx=10, sticky='nsew')
        self.label_docx_nmb = ttk.Label(self.folder_frame, text='.docx文件数')
        self.label_docx_nmb.config(font=('Microsoft YaHei UI', 10))
        self.label_docx_nmb.configure(justify="center", anchor="e")
        self.label_docx_nmb.grid(row=0, column=0, padx=15, pady=5, sticky="w")
        self.label_doc_nmb = ttk.Label(self.folder_frame, text='.doc文件数')
        self.label_doc_nmb.config(font=('Microsoft YaHei UI', 10))
        self.label_doc_nmb.configure(justify="center", anchor="e")
        self.label_doc_nmb.grid(row=0, column=1, padx=15, pady=5, sticky="w")
    def select_path_frame(self):
        mnplt_frame = ttk.LabelFrame(self, text="文件目录")
        mnplt_frame.grid(row=5, column=0, columnspan=2, padx=10, pady=10, ipadx=10, sticky='nsew')

        self.entry_path = ttk.Entry(mnplt_frame, justify=tk.LEFT, width=62,
                                    font=('Microsoft YaHei UI', 11))
        self.entry_path.insert(0, "浏览并选择目录")
        # self.entry_path.bind("<FocusIn>", lambda e: self.entry_path.delete('0', 'end'))
        # self.entry_path.focus_force()
        self.entry_path.grid(row=0, column=0, columnspan=2, padx=10, pady=10, sticky='ew')

        self.browse_button = tk.Button(mnplt_frame, text="选择目录", command=self.select_directory, bg='grey',
                                       fg='white',
                                       font=('Microsoft YaHei UI', 11, 'normal'))
        self.browse_button.grid(row=1, column=0, padx=10, pady=5, ipadx=10, ipady=5)

        self.start_remove_button = tk.Button(mnplt_frame, text="开始处理", command=self.start_remove_strings_from_files,
                                             bg='purple',
                                             fg='white',
                                             font=('Microsoft YaHei UI', 11, 'bold'))
        self.start_remove_button.grid(row=1, column=1, padx=10, pady=5, ipadx=10, ipady=5)

        self.read_all_item_to_list_box()

        geekerWin.protocol("WM_DELETE_WINDOW", self.on_window_close)

    def author_frame(self):
        author_frame = ttk.LabelFrame(self, text="关于")
        # author_frame = ttk.Frame(self)
        author_frame.grid(row=6, column=0, columnspan=2, padx=10, pady=10, ipadx=10)

        label_link = ttk.Label(author_frame, text='www.ITGeeker.net', font=('Microsoft YaHei UI', 10), cursor="hand2")
        label_link.bind("<Button-1>", lambda e: self.open_website("https://www.itgeeker.net"))
        label_link.grid(row=0, column=0, padx=(10, 0), ipadx=10, ipady=5, sticky="w")

        label_ver = ttk.Label(author_frame, text='开源版本Ver 1.1.2.0', font=('Microsoft YaHei UI', 10), cursor="heart")
        label_ver.config(font=('Microsoft YaHei UI', 10))
        label_ver.bind("<Button-1>",
                       lambda e: self.open_website(
                           "https://www.itgeeker.net/itgeeker-technical-service/itgeeker_convert_doc_to_docx/"))
        label_ver.grid(row=0, column=1, padx=(10, 10), ipadx=10, ipady=5, sticky="e")


if __name__ == "__main__":
    icon_b64 = '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'
    icondata = base64.b64decode(icon_b64)
    tmp_p = tempfile.gettempdir()
    tempFile = os.path.join(tmp_p, "icon.ico")
    iconfile = open(tempFile, "wb")
    iconfile.write(icondata)
    iconfile.close()

    geekerWin = tk.Tk()
    geekerWin.wm_iconbitmap(tempFile)
    ## Delete the tempfile
    # os.remove(tempFile)

    # geekerWin.geometry("500x580")
    # geekerWin.eval('tk::PlaceWindow . center')
    window_width = 638
    window_height = 730
    display_width = geekerWin.winfo_screenwidth()
    display_height = geekerWin.winfo_screenheight()
    left = int(display_width / 2 - window_width / 2)
    top = int(display_height / 2 - window_height / 2)
    geekerWin.geometry(f'{window_width}x{window_height}+{left}+{top}')

    geekerWin.title("技术奇客小工具-文档字符移除")

    list_sheet = ListSheet(geekerWin)
    list_sheet.pack()
    geekerWin.mainloop()
