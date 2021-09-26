# Topsis 模型数据处理模块文档



> author: 陈坤
>
> date: 2021.8
## 模块功能

数据处理: **数据的向量规范化, 数据的加权, 求解正理想解与负理想解。**




## 模块DataProcess输入参数
```python
    def __init__(self, data: np.array, w: np.array):
```
输入参数:

param: **data**  *输入数据数组(每一列为一个指标) -> np.array*

param: **w**  *权重系数 -> np.array*

​        

## 模块属性及总接口

---

|属性|
|---|
|DataProcess.data 输入数据数组 -> np.array|
|DataProcess.w 输入权重 -> np.array|
|DataProcess.m, DataProcess.n 输入数据数组行数与列数 -> int|
|DataProcess.Cstar 正理想解 -> array|
|DataProcess.C0 负理想解 -> array|

|总调用接口|
|---|
|DataProcess.work()|



## 使用示例

```python
if __name__=='__main__':
    testData = np.array([[3.0, 375, 3],
                         [6, 400, 5],
                         [8, 280, 3],
                         [2, 425, 2],
                         [1, 500, 4]])
    testW = np.array([0.2, 0.5, 0.3])
    dataProcess = DataProcess(testData, testW)
    dataProcess.work()
    print(dataProcess.Cstar, '\n', dataProcess.C0)

```
***打印结果:***
```python
[0.14985373 0.27783781 0.18898224] 
[0.01873172 0.15558917 0.07559289]
```
