"""
总结各学科所属项目的总金额,平均得分,团队数目
author: CoderChen
"""
import os
import xlrd
import xlwt
from xlutils.copy import copy


def getFileName(fileWalk):
    """
    获得目录下所有文件的文件名,不包括此Python文件
    :return: fileList 文件名列表
    """
    for path, b, fileName in os.walk(fileWalk):
        if path == 'C:\\Users\\CK\\Desktop\\temp':
            return fileName


class SubjectAnalyse:
    def __init__(self):
        self.avScore = None
        self.totalFunding = None
        self.totalTeam = None
        self.index = 0

    def ReadExcel(self, excel_name, sheet_name):
        excel = xlrd.open_workbook(excel_name)
        sheet = excel.sheet_by_name(sheet_name)
        self.totalFunding = 0
        totalScore = 0
        for i in range(1, sheet.nrows):
            funding = sheet.cell_value(i, 3)
            numProject = sheet.cell_value(i, 4)
            score = sheet.cell_value(i, 5)
            self.totalFunding += funding * numProject
            totalScore += score
        self.totalTeam = sheet.nrows - 1
        self.avScore = totalScore / self.totalTeam

    def WriteExcel(self, subject):
        # 目录中已存在
        excel_xlrd = xlrd.open_workbook('Table.xls', formatting_info=True)
        excel_xlwt = copy(wb=excel_xlrd)
        sheet_xlwt = excel_xlwt.get_sheet(0)
        if self.index == 0:
            title = ['学科', '团队数目', '总金额', '平均综合得分']
            for i in range(4):
                sheet_xlwt.write(0, i, title[i])
            self.index += 1

        row_value = [subject, self.totalTeam, self.totalFunding, self.avScore]
        for i in range(4):
            sheet_xlwt.write(self.index, i, row_value[i])
        self.index += 1
        excel_xlwt.save('Table.xls')


    def work(self, fileNames):
        for name in fileNames:
            subject = name.replace('.xls', '')
            self.ReadExcel(name, subject)
            self.WriteExcel(subject)
            print(name + '文件已完成')


if __name__ == '__main__':
    fileWalk = "C:\\Users\\CK\\Desktop\\temp"
    names = getFileName(fileWalk)
    names.remove('main.py')
    names.remove('Table.xls')
    subjectAnalyse = SubjectAnalyse()
    subjectAnalyse.work(names)

