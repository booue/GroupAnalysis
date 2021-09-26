"""
知网爬取项目关键词使初步划分学科
author: CoderChen
"""

import requests
import json
import xlrd
from xlutils.copy import copy
import time


class CrawlerCnki:
    """
    爬取知网某个关键词下所有文献的科目归属
    """
    def __init__(self, keyword: str):
        """
        类的初始化
        :param keyword: 写入的关键词
        """
        self.keyword = keyword
        self.url = 'https://kns.cnki.net/KNS8/Visual/GetGroupData'
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.111 Safari/537.36',
        }
        self.subjects = []
        self.amount = []
        self.field = []
        self.weight = []
        self.data = []

    def get_subject(self, num: int):
        """
        具体获取学科列表
        :param num: 需要保存前几名学科的学科数目
        :return:
        """
        try:
            data = {
                'dbCode': 'SCDB',
                'queryJson': '{"Platform":"","DBCode":"SCDB","KuaKuCode":"CJFQ,CDMD,CIPD,CCND,CISD,SNAD,BDZK,CCJD,CCVD,CJFN,CCJD","QNode":{"QGroup":[{"Key":"Subject","Title":"","Logic":1,"Items":[{"Title":"主题","Name":"SU","Value":"' + self.keyword + '","Operate":"%=","BlurType":""}],"ChildItems":[]}]}}',
                'groupName': '5-学科',
                'groupCode': 'SUBJECT',
                'valueType': '3'
            }
            requests_text = requests.post(url=self.url, headers=self.headers, data=data).text
            requests_json = json.loads(requests_text)
            for i in range(num):
                # 学科
                self.subjects.append(requests_json[i]['name'])
                # 学科文献数量
                self.amount.append(requests_json[i]['y'])
                # 学科代码
                self.field.append(requests_json[i]['c_fieldValue'])

            # 计算学科文献比重
            amount_sum = sum(self.amount)
            for i in self.amount:
                self.weight.append(round(i / amount_sum, 2))

            # 将所有信息纳入data列表
            for i in range(num):
                self.data.append(self.subjects[i])
                self.data.append(self.field[i])
                self.data.append(self.weight[i])

        except json.decoder.JSONDecodeError:
            try:
                # 在知网中"以主题方式搜索"不含有该关键词的任何文献,则改变为"以全文方式搜索"
                data_plus = {
                    'dbCode': 'SCDB',
                    'queryJson': '{"Platform":"","DBCode":"SCDB","KuaKuCode":"CJFQ,CDMD,CIPD,CCND,CISD,SNAD,BDZK,CCJD,CCVD,CJFN,CCJD","QNode":{"QGroup":[{"Key":"Subject","Title":"","Logic":1,"Items":[{"Title":"全文","Name":"SU","Value":"' + self.keyword + '","Operate":"%=","BlurType":""}],"ChildItems":[]}]}}',
                    'groupName': '5-学科',
                    'groupCode': 'SUBJECT',
                    'valueType': '3'
                }
                requests_text = requests.post(url=self.url, headers=self.headers, data=data_plus).text
                requests_json = json.loads(requests_text)
                for i in range(num):
                    self.subjects.append(requests_json[i]['name'])
                    self.amount.append(requests_json[i]['y'])
                    self.field.append(requests_json[i]['c_fieldValue'])
                amount_sum = sum(self.amount)
                for i in self.amount:
                    self.weight.append(round(i / amount_sum, 2))
                for i in range(num):
                    self.data.append(self.subjects[i])
                    self.data.append(self.field[i])
                    self.data.append(self.weight[i])

            except json.decoder.JSONDecodeError:
                # 若全文搜索都未检索到任何文献,则删减关键词重新检索
                self.keyword = ' '.join(self.keyword.split()[-1])
                try:
                    # 在知网中"以主题方式搜索"不含有该关键词的任何文献,则改变为"以全文方式搜索"
                    data_plus = {
                        'dbCode': 'SCDB',
                        'queryJson': '{"Platform":"","DBCode":"SCDB","KuaKuCode":"CJFQ,CDMD,CIPD,CCND,CISD,SNAD,BDZK,CCJD,CCVD,CJFN,CCJD","QNode":{"QGroup":[{"Key":"Subject","Title":"","Logic":1,"Items":[{"Title":"全文","Name":"SU","Value":"' + self.keyword + '","Operate":"%=","BlurType":""}],"ChildItems":[]}]}}',
                        'groupName': '5-学科',
                        'groupCode': 'SUBJECT',
                        'valueType': '3'
                    }
                    self.requests_text = requests.post(url=self.url, headers=self.headers, data=data_plus).text
                    requests_json = json.loads(self.requests_text)
                    for i in range(num):
                        self.subjects.append(requests_json[i]['name'])
                        self.amount.append(requests_json[i]['y'])
                        self.field.append(requests_json[i]['c_fieldValue'])
                    amount_sum = sum(self.amount)
                    for i in self.amount:
                        self.weight.append(round(i / amount_sum, 2))
                    for i in range(num):
                        self.data.append(self.subjects[i])
                        self.data.append(self.field[i])
                        self.data.append(self.weight[i])

                except IndexError:
                    # 所含学科少于num
                    amount_sum = sum(self.amount)
                    for i in self.amount:
                        self.weight.append(round(i / amount_sum, 2))
                    for i in range(len(self.subjects)):
                        self.data.append(self.subjects[i])
                        self.data.append(self.field[i])
                        self.data.append(self.weight[i])

                except json.decoder.JSONDecodeError:
                    self.subjects = []

            except IndexError:
                # 所含学科少于num
                amount_sum = sum(self.amount)
                for i in self.amount:
                    self.weight.append(round(i / amount_sum, 2))
                for i in range(len(self.subjects)):
                    self.data.append(self.subjects[i])
                    self.data.append(self.field[i])
                    self.data.append(self.weight[i])

        except IndexError:
            # 所含学科少于num
            amount_sum = sum(self.amount)
            for i in self.amount:
                self.weight.append(round(i / amount_sum, 2))
            for i in range(len(self.subjects)):
                self.data.append(self.subjects[i])
                self.data.append(self.field[i])
                self.data.append(self.weight[i])


def writesheet(row_values: list, row: int, filename, col_begin=0):
    """
    写入表单某一列数据
    :param: row_values: 写入的数据列表
    :param row: 写到第几行
    :param filename: 写入的文件名
    :param col_begin: 写入的起始列数,默认为 0
    :return: 保存文件
    """
    data = xlrd.open_workbook(filename, formatting_info=True)
    excel = copy(wb=data)
    excel_table = excel.get_sheet(0)
    for i in range(col_begin, col_begin + len(row_values)):
        excel_table.write(row, i, row_values[i-col_begin])  # 参数为行,列,值
        excel.save(filename)


if __name__ == '__main__':
    # cnki = CrawlerCnki('高效 频率')
    # cnki.get_subject(10)
    # print(cnki.data)

    excel = xlrd.open_workbook('project_output_breadth_result.xls')
    sheet = excel.sheets()[0]
    for i in range(750, sheet.nrows):

        try:
            temp_value = sheet.cell_value(i, 3)
            if temp_value == '':
                cell = sheet.cell(i, 2).value
                cnki = CrawlerCnki(cell)

                cnki.get_subject(10)
                writesheet(cnki.data, i, "project_output_breadth_result.xls", col_begin=3)
                print("第" + str(i) + "行数据已保存!   剩余" + str(sheet.nrows - i) + "行数据")
                time.sleep(1)

        except PermissionError:
            cell = sheet.cell(i, 2).value
            cnki = CrawlerCnki(cell)

            cnki.get_subject(10)
            writesheet(cnki.data, i, "project_output_breadth_result.xls", col_begin=3)
            print("第" + str(i) + "行数据已保存!   剩余" + str(sheet.nrows - i) + "行数据")
            time.sleep(1)





