# 相关包引入
from fuzzywuzzy import fuzz
import jieba.analyse
import pandas as pd


# 停用词汇读取
with open("stop_word.txt","r",encoding = "utf-8") as f:
    content = f.readlines()
stop_words = []
for ones in content:
    stop_words.append(ones.strip('\n'))

# 专业词汇读取
with open("specialized  vocabulary.txt","r",encoding = "utf-8") as f:
    content = f.readlines()
specialized_words = []
for ones in content:
    specialized_words.append(ones.strip('\n'))


def matching_content(content):
    '''
    :param content:待匹配文本 
    :return: 匹配程度高的专业词汇(specaiized word)
    '''
    field = {}
    spe_word = ""
    for ones in specialized_words:
        degree = fuzz.ratio(content,ones)
        if degree >= 25:
            field[degree] = ones
        else:
            continue
    if field != {}:
        spe_word = field[max(field.keys())]
    else:
        spe_word = ""
    return spe_word


def stop_content(lst):
    '''
    function---（1）停用词的去除（2）专业词汇添加
    :param content: 待处理文本
    :return: 处理后文本
    '''
    lst_deal = []
    for ones in lst:
        if ones not in stop_words:
            lst_deal.append(ones)
        else:
            continue
    return lst_deal


def text_dealing(content):
    '''
    :param content:待匹配文本 
    :return: 处理后的关键词文本
    '''
    text = jieba.analyse.extract_tags(content,topK = 3)
    text_pre1 = stop_content(text)
    text_pre2 = matching_content(content)
    text_pre1.append(text_pre2)
    return text_pre1


# 此处修改文件路径，分别读取横向，纵向，外协文件
data = pd.read_csv("C:\python_learning\Group analysis\project_breadth.csv",encoding = "gbk")
data_apply = data["项目名称"]


project_apply = []
for ones in data_apply:
    project_apply.append(ones)

content_apply = []
for ones in data_apply:
    ones_after = text_dealing(ones)
    str = ""
    for one in ones_after:
        str = str + " " + one
    content_apply.append(str)


dataframe = pd.DataFrame({"项目原名称":project_apply,"项目名称缩写":content_apply})
# 此处修改输出文件路径，将提取到的关键词写入对应的文件中
dataframe.to_csv("project_output_breadth.csv",index = True,sep = ",")














