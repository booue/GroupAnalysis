import pandas as pd
from pandas.core.frame import DataFrame

def time_interval(worksheet_name,origin_sheet_name):
    data=pd.read_excel(r"项目YJ-01--开发用途（内部使用）.xlsx",sheet_name=origin_sheet_name)
    project_name=data["项目名称"]
    write_data=pd.read_excel(r"年均经费（项目版）.xlsx",sheet_name=worksheet_name)
    write_data["合同金额"]=data["合同金额"]
    write_data["项目名称"]=data["项目名称"]
    write_data["立项时间"]=data["立项时间"]
    write_data["负责人工号"]=data["负责人工号"]
    write_data["计划完成日期"]=data["计划完成日期"]
    write_data["负责人工号"]=data["负责人工号"]
    write_data["项目历时"]=pd.to_datetime(write_data["计划完成日期"]) -pd.to_datetime(write_data["立项时间"])
    temp_days=[0 for x in range(len(project_name))]
    temp_days=write_data["项目历时"].map(lambda x:x.days)
    for i in range(len(project_name)):
        if temp_days[i]>=0 and temp_days[i]<=365:
            write_data["项目历时"][i]=365
        else:
            write_data["项目历时"][i]=temp_days[i]
    for i in range(len(project_name)):
       write_data["项目来源"][i]=origin_sheet_name
       write_data["年均经费"][i]=write_data["合同金额"][i]/(float(write_data["项目历时"][i]+1)/365)
    return write_data

def merge1(df1,df2,df3):
    df4=pd.concat([df1,df2])
    df4=pd.concat([df3,df4])
    return df4

if __name__=="__main__":
    breadth=time_interval("breadth","横")
    length=time_interval("length","纵")
    external=time_interval("external","外协")
    breadth_result=(
        breadth.groupby(by=breadth["负责人工号"])
               .agg(
                   项目名称=("项目名称",lambda x:",".join(x)),
                   合同金额=("合同金额","sum"),
                   年均经费=("年均经费","sum"),
                   项目来源=("项目来源",lambda x:",".join(x.unique())),
               )
               .reset_index()
               .rename(columns={"负责人工号":"负责人工号"})
    )
    breadth_result["项目数量"]=breadth_result["项目名称"].apply(lambda x:len(x.split(",")))
    breath_result=breadth_result.sort_values(by="负责人工号")

    length_result=(
        length.groupby(by=length["负责人工号"])
               .agg(
                   项目名称=("项目名称",lambda x:",".join(x)),
                   合同金额=("合同金额","sum"),
                   年均经费=("年均经费","sum"),
                   项目来源=("项目来源",lambda x:",".join(x.unique())),
               )
               .reset_index()
               .rename(columns={"负责人工号":"负责人工号"})
    )
    length_result["项目数量"]=length_result["项目名称"].apply(lambda x:len(x.split(",")))
    length_result=length_result.sort_values(by="负责人工号")
    
    external_result=(
        external.groupby(by=external["负责人工号"])
               .agg(
                   项目名称=("项目名称",lambda x:",".join(x)),
                   合同金额=("合同金额","sum"),
                   年均经费=("年均经费","sum"),
                   项目来源=("项目来源",lambda x:",".join(x.unique())),
               )
               .reset_index()
               .rename(columns={"负责人工号":"负责人工号"})
    )
    external_result["项目数量"]=external_result["项目名称"].apply(lambda x:len(x.split(",")))
    external_result=external_result.sort_values(by="负责人工号")

    all_result=merge1(breadth_result,length_result,external_result)
    all_result=(
        all_result.groupby(by=all_result["负责人工号"])
               .agg(
                   项目名称=("项目名称",lambda x:",".join(x)),
                   合同金额=("合同金额","sum"),
                   年均经费=("年均经费","sum"),
                   项目来源=("项目来源",lambda x:",".join(x.unique())),
               )
               .reset_index()
               .rename(columns={"负责人工号":"负责人工号"})
    )
    all_result["项目数量"]=all_result["项目名称"].apply(lambda x:len(x.split(",")))
    all_result=all_result.sort_values(by="负责人工号")

    with pd.ExcelWriter(r"年均经费（负责人版）.xlsx") as writer:
        breath_result.to_excel(writer,sheet_name="breadth")
        length_result.to_excel(writer,sheet_name="length")
        external_result.to_excel(writer,sheet_name="external")
        all_result.to_excel(writer,sheet_name="breadth_length_external")


