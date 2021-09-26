"""
将表格按照所属学科划分到不同表格中
author: CoderChen
"""
import xlrd
import xlwt
from xlutils.copy import copy


def WriteExcel(subject, row_value):
    excel_name = subject + '.xls'
    try:
        # 目录中已存在该表
        excel_xlrd = xlrd.open_workbook(excel_name, formatting_info=True)
        sheet_xlrd = excel_xlrd.sheets()[0]
        excel_xlwt = copy(wb=excel_xlrd)  # 完成 xlrd 对象向 xlwt 对象的转变
        sheet_xlwt = excel_xlwt.get_sheet(0)  # 获得操作的页
        for i in range(4):
            sheet_xlwt.write(sheet_xlrd.nrows, i, row_value[i])
        excel_xlwt.save(excel_name)

    except FileNotFoundError:
        # 若目录中不存在此表,则创建新表格
        excel = xlwt.Workbook()
        sheet = excel.add_sheet(subject)
        title = ['项目名称', '负责人工号', '跨学科数', '年均经费']
        for i in range(4):
            sheet.write(0, i, title[i])
        for i in range(4):
            sheet.write(1, i, row_value[i])
        excel.save(excel_name)


def ReadExcel(excel_name):
    TotalTable = xlrd.open_workbook(excel_name)
    sheet = TotalTable.sheet_by_name('breadth_length_external')
    for i in range(1, sheet.nrows):
        subjects = sheet.cell_value(i, 4)
        subject_list = subjects.split(',')
        row_value = []
        for j in [1, 2, 3, 5]:
            row_value.append(sheet.cell(i, j).value)
        work(subject_list, row_value)
        print("已完成" + str(i) + "行,剩余" + str(sheet.nrows - i) + "行")


def work(subjects_list, row_value):
    for subject in subjects_list:
        WriteExcel(subject, row_value)


if __name__ == '__main__':
    excel_name = 'TotalTable.xls'
    ReadExcel(excel_name)
