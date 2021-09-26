### 二次匹配学科确定

+++++++++++++++

##### 1.“二次匹配学科确定”项目模块说明：

+ 目的：知网爬虫获得的学科列表借助于承办单位进行二次匹配，建模分析获得项目对应最终学科
+ 步骤：
  + 知网爬虫学科二次匹配
  + “类决策树”方法确定学科归属（具体分析见《科研团队量化_建模文档》）

##### 2.文件说明：

+ organizer_subject_rate.py          知网爬虫学科借助承办单位二次匹配
+ final_subject_result.py          “类决策树”方法确定项目学科归属
+ 学科归属结果文件：
  + breadth_organizer_ratio_result.xls          横向项目学科归属结果文件
  + external_organizer_ratio_result.xls          外协项目学科归属结果文件
  + length_organizer_ratio_result.xls          纵向项目学科归属结果文件
