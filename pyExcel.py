import pandas as libPandas

# 使用ExcelFile类加载Excel文件
tXls = libPandas.ExcelFile('example.xls')

# 获取所有工作表的名称
tSheetNames = tXls.sheet_names

# 打印工作表名称
print(tSheetNames)