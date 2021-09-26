from numpy import nan
import pandas as pd
from pandas import DataFrame
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

def string_similar(list1,s2):
    subject_ratio=process.extractOne(s2,list1)
    return subject_ratio[1]

def operate_sheet(workbook_name,worksheet_name,origin_sheet_name,write_workbook_name,write_worksheet_name):
    data_origin_sheet=pd.read_excel(r"项目YJ-01--开发用途（内部使用）.xlsx",sheet_name=origin_sheet_name)
    organizer_info=data_origin_sheet["承担单位"] 
    info_length=len(organizer_info)
    data_processed=pd.read_excel(workbook_name,sheet_name=worksheet_name)
    project_name=data_processed["项目原名称"]
    project_name_processed=data_processed["项目名称缩写"]
    cnki_subject=[ " " for x in range(10*info_length)]
    data_organizer_subject=pd.read_excel(r"学院对应专业及学科.xlsx",sheet_name="Sheet1",nrows=11)
    write_data=pd.read_excel(write_workbook_name,sheet_name=write_worksheet_name,nrows=3*len(organizer_info)+1)
    for i in range(info_length):
        #项目名称获取
        write_data["项目原名称"][3*i]=project_name[i]
        write_data["项目名称缩写"][3*i]=project_name_processed[i]
        #承办单位的获取
        write_data["承担单位"][3*i]=organizer_info[i]
        #单元格标识
        write_data["匹配度相关"][3*i]="知网前十学科"
        write_data["匹配度相关"][3*i+1]="承担单位匹配度"
        write_data["匹配度相关"][3*i+2]="最终结果"
        #学科数目获取，以便后续归一化
        subject_num=10
        while subject_num>0:
            if str(data_processed["学科"+str(subject_num)][i])!='nan':
                break
            else:
                subject_num=subject_num-1
        organizer_subject=[" " for x in range(subject_num)]
        #知网前十专业获取 
        for j in range(subject_num):
            cnki_subject[i*10+j]=data_processed["学科"+str(j+1)][i]
            write_data["学科"+str(j+1)][3*i]=cnki_subject[i*10+j]
            write_data["学科"+str(j+1)+"代号"][3*i]=data_processed["学科"+str(j+1)+"代号"][i]
            write_data["学科"+str(j+1)+"比率"][3*i]=data_processed["学科"+str(j+1)+"比率"][i]
        #学院对应学科的获取
        for j in range(subject_num):
            organizer_subject[j]=data_organizer_subject[organizer_info[i]][j]
        #匹配度获取
        ratio=[0 for x in range(subject_num)]
        for j in range(1,subject_num+1):
            ratio[j-1]=string_similar(organizer_subject,str(write_data["学科"+str(j)][3*i]))
        sum_ratio=sum(ratio)  
        for j in range(1,subject_num+1):
            if sum_ratio: 
                write_data["学科"+str(j)+"比率"][3*i+1]="%.2f" %(ratio[j-1]/sum_ratio)
            else:
                write_data["学科"+str(j)+"比率"][3*i+1]="%.2f" %(ratio[j-1])  
        
        
    return write_data
    
if __name__=="__main__":
    df_write_breadth=operate_sheet("project_output_breadth_result.xls","project_output_breadth","横","breadth_organizer_ratio_result.xls","project_output_breadth")
    df_write_length=operate_sheet("project_output_length_result.xls","project_output_length","纵","length_organizer_ratio_result.xls","project_output_length")
    df_write_external=operate_sheet("project_output_external_result.xls","project_output_external","外协","external_organizer_ratio_result.xls","project_output_external")
    DataFrame(df_write_breadth).to_excel("breadth_organizer_ratio_result.xls",sheet_name="project_output_breadth")
    DataFrame(df_write_length).to_excel("length_organizer_ratio_result.xls",sheet_name="project_output_length")
    DataFrame(df_write_external).to_excel("external_organizer_ratio_result.xls",sheet_name="project_output_external")

    
