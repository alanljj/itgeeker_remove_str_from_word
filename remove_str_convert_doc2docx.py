# -*- coding: utf-8 -*-
import glob
import os
import win32com.client


def separate_file_info_by_ffp(ffp):
    dirname = os.path.dirname(ffp)
    filename, file_extension = os.path.splitext(ffp)
    # basename = os.path.basename(ffp)
    basename_no_ext = os.path.splitext(os.path.basename(ffp))[0]
    print(basename_no_ext)
    print('dirname: ', dirname)
    # print('filename: ', filename)
    print('file_extension: ', file_extension)
    # print('basename: ', basename)
    print('basename_no_ext: ', basename_no_ext)
    return dirname, basename_no_ext, file_extension


def convert_doc2docx_by_win32com(doc_list):
    word = win32com.client.Dispatch("Word.Application")
    word.visible = 0
    doc2docx_l = []

    # for i, doc in enumerate(glob.iglob("*.doc")):
    # for i, doc in enumerate(glob.iglob(r"D:\tmp\hrefond_tmplt\*.doc")):
    # for i, doc in enumerate(glob.iglob(f_path + r"\*.doc")):
    for i, doc in enumerate(doc_list):
        in_file = os.path.abspath(doc)
        # print('in_file: %s' % in_file)
        wb = word.Documents.Open(in_file)

        dirname, basename_no_ext, file_extension = separate_file_info_by_ffp(in_file)
        new_docx_f = os.path.join(dirname, basename_no_ext + '-converted' + '.docx')
        # print('new_docx_f: %s' % new_docx_f)
        doc2docx_l.append(new_docx_f)
        # out_file = os.path.abspath("out{}.docx".format(i))
        # wb.SaveAs2(out_file, FileFormat=16)  # file format for docx

        wb.SaveAs2(new_docx_f, FileFormat=16)  # file format for docx
        wb.Close()
    word.Quit()

    print('doc2docx_l: %s' % doc2docx_l)
    return doc2docx_l
