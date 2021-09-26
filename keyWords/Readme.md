### 关键词提取-keyWords_dealing

+++++++

##### 1.“关键词提取”项目模块说明：

+ 目的：项目全名称处理，保证后续项目名称输入至知网检索可获得结果
+ 步骤：
  + 结巴中文分析模块处理
  + 自定义停用词库去除停用词
  + 自定义专业词汇库添加专业词汇

##### 2.文件说明：

+ keyWords_dealing.py          关键词提取源代码文件
+ stop_word.txt          停用词文件（代码辅助文件）
+ specialized vocabulary.txt          专用词汇文件（代码辅助文件）
+ 结果文件
  + project_breath.csv          横向关键词提取结果文件
  + project_length.csv          纵向关键词提取结果文件
  + project_external.csv          外协关键词提取结果文件