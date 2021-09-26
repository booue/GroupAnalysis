import numpy as np
import pandas as pd
from numpy import matlib
import xlwt as xw
from xlutils.copy import copy
import xlrd as xr
import os




class DataProcess:
    """
    数据处理: 数据的向量规范化, 数据的加权, 求解正理想解与负理想解
    author: CK
    """
    def __init__(self, data: np.array, w: np.array):
        """
        类的初始化
        :param data: 输入数据数组(每一列为一个指标) -> array
        :param w: 权重系数 -> array
        """
        self.data = data
        self.w = w
        self.m, self.n = self.data.shape
        self.Cstar = []  # 正理想解
        self.C0 = []  # 负理想解

    def __Standard(self):
        """
        数据的向量标准化
        """
        K = np.power(np.sum(pow(self.data, 2), axis=0), 0.5)
        for i in range(len(K)):
            self.data[:, i] = self.data[:, i] / K[i]

    def __Weighted(self):
        """
        向量标准化后的数据进行加权处理
        """
        converted_w = np.tile(self.w, (self.m, 1))
        self.data = np.multiply(self.data, converted_w)

    def __SolveBestAndLeast(self):
        """
        求解正理想解和负理想解
        """
        self.Cstar = np.amax(self.data, axis=0)
        self.C0 = np.amin(self.data, axis=0)

    def work(self):
        self.__Standard()
        self.__Weighted()
        self.__SolveBestAndLeast()




def read_processed_alldata():
    '''
    read_processed_alldata->总项目
    评价指标:一列：学科数量  二列：项目数量  三列：年均经费
    :return:dataArray->数据读取返回列表 
    '''
    data=pd.read_excel(r"团队排名（UESTC）.xls")
    level1_subject_num=data["学科数量"]
    project_num=data["项目数量"]
    annual_funding=data["年均经费"]
    data_length=len(data["负责人工号"])
    data_array=matlib.zeros((data_length,3))

    data_array[:,[0]]=np.array([list(level1_subject_num)]).T
    data_array[:,[1]]=np.array([list(project_num)]).T
    data_array[:,[2]]=np.array([list(annual_funding)]).T
    return data_array


def distanceMaxMin(dataArray,inx):
    '''
    :param dataArray:原数据矩阵
    :param inx: 目标矩阵
    :return: 距离列表
    '''
    size = dataArray.shape[0]
    diffMat = np.tile(inx, (size, 1)) - dataArray
    spdiffMat = diffMat ** 2
    spdisances = spdiffMat.sum(axis=1) ** 0.5
    return spdisances


def assesment(dataArray,wList):
    '''
    :param dataArray:数据矩阵 
    :param wList: 权重矩阵
    :return: 综合指数
    '''

    # 数据预处理
    testData = np.array(dataArray)
    testW = np.array(wList)
    dataProcess = DataProcess(testData, testW)
    dataProcess.work()

    dataApply = dataProcess.data
    max = dataProcess.Cstar # 理想最优解
    min = dataProcess.C0    # 理想最劣解

    maxDistances = distanceMaxMin(dataApply,max)
    minDistances = distanceMaxMin(dataApply,min)

    result = minDistances/(maxDistances + minDistances).tolist() # 评价指标建立
    return result



# 综合排名
dataArray = read_processed_alldata()
wList = [0.2,0.3,0.5]
result = assesment(dataArray,wList)

file="团队排名（UESTC）.xls"
style = xw.easyxf()
oldwb = xr.open_workbook(file) # 工作薄读取
newwb = copy(oldwb) # 新工作薄复制
newws = newwb.get_sheet(0) # 获取指定工作表
for i in range(len(result)):
    newws.write(i+1,7, result[i],style) # 内容写入特定列
newwb.save(file) # 修改保存















