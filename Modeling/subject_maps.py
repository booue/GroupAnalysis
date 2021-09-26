import pandas as pd
import re

def map_relation():
    maps=pd.read_excel("映射表.xls",sheet_name="映射表")
    cnki_subject=maps['知网学科'].apply(lambda x:''.join(re.findall('[\u4e00-\u9fa5,\u3001]',x)))
    standard_subject=maps['标准学科'].apply(lambda x:",".join(''.join([i for i in x if not i.isdigit()]).split("/")).strip(" "))
    map_dict=dict(zip(cnki_subject,standard_subject))
    return map_dict

def sheet_map(workbook_name,worksheet_name):
    map_dict=map_relation()
    print(map_dict)
    data=pd.read_excel(workbook_name,sheet_name=worksheet_name)
    processed_subject_list=[ " "for x in range(len(data["项目名称"]))]
    for i in range(len(data['项目名称'])):
        original_subject_list=data['总学科内容'][i].split(",")
        processed_subject=[]
        for j in range(len(original_subject_list)):
            processed_subject.append(map_dict[original_subject_list[j]])
        processed_subject_list[i]=",".join(list({}.fromkeys(processed_subject).keys()))
        data['映射后学科内容'][i]=processed_subject_list[i]
        data['总跨学科数目'][i]=len(data['映射后学科内容'][i].split(','))
    return data        



if __name__=="__main__":
    breadth_result=sheet_map("学科实力评估（项目版）.xlsx",'breadth')
    length_result=sheet_map("学科实力评估（项目版）.xlsx",'length')
    external_result=sheet_map("学科实力评估（项目版）.xlsx",'external')
    all_result=sheet_map("学科实力评估（项目版）.xlsx",'breadth_length_external')

    with pd.ExcelWriter(r"学科实力评估（映射后）.xlsx") as writer:
        breadth_result.to_excel(writer,sheet_name="breadth")
        length_result.to_excel(writer,sheet_name="length")
        external_result.to_excel(writer,sheet_name="external")
        all_result.to_excel(writer,sheet_name="breadth_length_external")

