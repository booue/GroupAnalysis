import pandas as pd
from pandas import DataFrame
import numpy as np

def sheet_result(workboot_name,worksheet_name):
    data=pd.read_excel(workboot_name,sheet_name=worksheet_name)
    project_num=int(len(data["匹配度相关"])/3)
    for i in range(project_num):
        #获取学科数目
        subject_num=10
        while subject_num>0:
            if str(data["学科"+str(subject_num)][3*i])!='nan':
                break
            else:
                subject_num=subject_num-1
        print(subject_num)
        #根据学科数目开始进行处理
        if subject_num<=0:
            print("进入0学科判断")
            continue
        elif subject_num <=1:
            print("进入1学科判断")
            final_subject=data["学科1"][3*i]
            data["学科1"][3*i+2]=set([final_subject])
        elif subject_num <=2:
            print("进入2学科判断")
            result_1=0.4*float(data["学科1比率"][3*i])+0.6*float(data["学科1比率"][3*i+1])
            result_2=0.4*float(data["学科2比率"][3*i])+0.6*float(data["学科2比率"][3*i+1])
            if abs(result_1-result_2) <=0.1:
                final_subject=[data["学科1"][3*i],data["学科2"][3*i]]
            else:
                final_subject=[data["学科1"][3*i] if (result_1 >result_2) else data["学科2"][3*i]]
            data["学科1"][3*i+2]=set(final_subject)
        else:
            print("进入3学科判断")
            cnki_subject_order=[" " for x in range(subject_num)]
            cnki_subject_ratio=[0 for x in range(subject_num)]
            third_for_cnki=[" " for x in range(3)] #知网前三学科
            third_ratio_cnki=[0 for x in range(3)] #用以存储知网前三学科的比率
            organizer_subject_order=[" " for x in range(subject_num)]
            organizer_ratio=[0 for x in range(subject_num)]#原始承担单位比率
            organizer_subject_ratio=[0 for x in range(subject_num)]#排序后单位比率
            third_for_organizer=[" " for x in range(3)] #承担单位前三学科
            third_ratio_organizer=[0 for x in range(3)]
            #对知网和承担单位进行匹配度排序
            for j in range(subject_num):
                cnki_subject_order[j]=data["学科"+str(j+1)][3*i]
                cnki_subject_ratio[j]=data["学科"+str(j+1)+"比率"][3*i]
                organizer_ratio[j]=data["学科"+str(j+1)+"比率"][3*i+1]
            #判断是否所有学科匹配均为0
            if (np.array(organizer_ratio) ==0).all():
                final_subject=data["学科1"][3*i]
                data["学科1"][3*i+2]=set([final_subject])
                continue
            cnki_ratio_dict=dict(zip(cnki_subject_order,cnki_subject_ratio))
            organizer_ratio_dict=dict(zip(cnki_subject_order, organizer_ratio))
            dic1SortList = sorted(organizer_ratio_dict.items(),key = lambda x:x[1],reverse = True)
            #排序后的承担单位学科和匹配率列表
            for j in range(subject_num):
                organizer_subject_order[j]=dic1SortList[j][0]
                organizer_subject_ratio[j]=dic1SortList[j][1]
            for j in range(3):
                third_for_cnki[j]=cnki_subject_order[j]
                third_for_organizer[j]=organizer_subject_order[j]
            set_common=set(third_for_cnki).intersection(set(third_for_organizer))
            if len(set_common)>=2:
                final_subject=set_common
                data["学科1"][3*i+2]=final_subject
            else:
                #分别得到知网前三对应的承担单位匹配率和承担单位前三对应的知网匹配率
                for j in range(3):
                    third_ratio_cnki[j]=organizer_ratio_dict[third_for_cnki[j]] 
                    third_ratio_organizer[j]=cnki_ratio_dict[third_for_organizer[j]]
                index_1=organizer_subject_ratio.index(max(third_ratio_cnki))
                index_2=cnki_subject_ratio.index(max(third_ratio_organizer))
                set_1=[organizer_subject_order[index_1]]
                set_2=[cnki_subject_order[index_2]]
                final_subject=set(set_1).union(set(set_2))
                final_subject=set(final_subject).union(set_common)
                data["学科1"][3*i+2]=final_subject
    DataFrame(data).to_excel(workboot_name,sheet_name=worksheet_name)
            

if __name__=="__main__":
    sheet_result("breadth_organizer_ratio_result.xls","breadth")
    sheet_result("length_organizer_ratio_result.xls","length")
    sheet_result("external_organizer_ratio_result.xls","external")