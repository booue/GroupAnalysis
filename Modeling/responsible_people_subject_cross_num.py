import pandas as pd
from pandas.core.frame import DataFrame

def merge1(df1,df2,df3):
    df4=pd.concat([df1,df2])
    df4=pd.concat([df3,df4])
    return df4

def subject_num(workbook_name,worksheet_name):
    data=pd.read_excel(workbook_name,worksheet_name)
    '''
    for i in range(len(data["一级学科内容"])):
        print(i)
        data["一级学科内容"][i]=",".join(eval(str(data["一级学科内容"][i])))
    '''
    data["一级学科内容"]=data["一级学科内容"].apply(lambda x:",".join(eval(str(x))))
    data["总学科内容"]=data["总学科内容"].apply(lambda x:",".join(eval(str(x))))
    data_result=(
        data.groupby(by=data["负责人工号"])
               .agg(
                   项目名称=("项目名称",lambda x:",".join(x)),
                   一级学科内容=("一级学科内容",lambda x:",".join(x.unique())),
                   总学科内容=("总学科内容",lambda x:",".join(x.unique())),
               )
               .reset_index()
               .rename(columns={"负责人工号":"负责人工号"})
    )
    data_result["项目数量"]=data_result["项目名称"].apply(lambda x:len(x.split(",")))
    data_result["一级学科内容"]=data_result["一级学科内容"].apply(lambda x:",".join(list({}.fromkeys(x.split(",")).keys())))
    data_result["跨一级学科数目"]=data_result["一级学科内容"].apply(lambda x:len(x.split(",")))
    data_result["总学科内容"]=data_result["总学科内容"].apply(lambda x:",".join(list({}.fromkeys(x.split(",")).keys())))
    data_result["总跨学科数目"]=data_result["总学科内容"].apply(lambda x:len(x.split(",")))
    data_result=data_result.sort_values(by="负责人工号")
    return data_result

if __name__=="__main__":
    breadth_result=subject_num(r"跨学科数量(项目版).xlsx","breadth")
    length_result=subject_num(r"跨学科数量(项目版).xlsx","length")
    external_result=subject_num(r"跨学科数量(项目版).xlsx","external")
    all_result=merge1(breadth_result,length_result,external_result)
    all_result=(
        all_result.groupby(by=all_result["负责人工号"])
               .agg(
                   项目名称=("项目名称",lambda x:",".join(x)),
                   一级学科内容=("一级学科内容",lambda x:",".join(x.unique())),
                   总学科内容=("总学科内容",lambda x:",".join(x.unique())),
               )
               .reset_index()
               .rename(columns={"负责人工号":"负责人工号"})
    )
    
    all_result["项目数量"]=all_result["项目名称"].apply(lambda x:len(x.split(",")))
    all_result["一级学科内容"]=all_result["一级学科内容"].apply(lambda x:",".join(list({}.fromkeys(x.split(",")).keys())))
    all_result["跨一级学科数目"]=all_result["一级学科内容"].apply(lambda x:len(x.split(",")))
    all_result["总学科内容"]=all_result["总学科内容"].apply(lambda x:",".join(list({}.fromkeys(x.split(",")).keys())))
    all_result["总跨学科数目"]=all_result["总学科内容"].apply(lambda x:len(x.split(",")))
    all_result=all_result.sort_values(by="负责人工号")
    
    with pd.ExcelWriter(r"跨学科数量(负责人版).xlsx") as writer:
        DataFrame(breadth_result).to_excel(writer,sheet_name="breadth")
        DataFrame(length_result).to_excel(writer,sheet_name="length")
        DataFrame(external_result).to_excel(writer,sheet_name="external")
        DataFrame(all_result).to_excel(writer,sheet_name="breadth_length_external")
    