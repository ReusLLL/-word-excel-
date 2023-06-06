# -*- coding: utf-8 -*-
"""
Created on Thu Feb  9 19:47:31 2023

@author: 95223
"""



# 导入需要的模块
import time
from xlutils.copy import copy # 从xlutils导入copy
import xlrd
import xlwt
import os

# 指定目录和模板文件路径
path = 'C:/Users/User/Desktop/数据底座_mapping_V1.1'
formwork_file = 'C:/Users/User/Desktop/数据底座模型映射/数据底座模型映射文档-模板.xls'

# 获取目录下的文件名列表
file_list = os.listdir(path)

# 如果第一个文件名以"~$"开头，去除它
if file_list[0][:2] == '~$':
    file_list[0] = file_list[0][2:]
    
# 打开模板文件并保留格式，准备写入新的数据
formbook = xlrd.open_workbook(filename=formwork_file, formatting_info = True)
wtbook = copy(formbook)
ts = wtbook.get_sheet('字段属性盘点')

# 创建表名和字段名的样式
tableNameStyle = xlwt.XFStyle()   # 创建一个样式对象
tableNameFont = xlwt.Font()       # 创建一个字体格式
tableNameFont.name = '宋体'   # 设置字体为微软雅黑
tableNameFont.height = 20*11      # 设置字号为11
tableNameStyle.font = tableNameFont

fieldNameStyle = xlwt.XFStyle()   # 创建一个样式对象
fieldNameFont = xlwt.Font()       # 创建一个字体格式
fieldNameFont.name = '微软雅黑'       # 设置字体为宋体
fieldNameFont.height = 20*11      # 设置字号为11
fieldNameStyle.font = fieldNameFont

start_r = 2 # 从第二行开始写入数据
for i in file_list:
    # 打开数据文件
    workbook = xlrd.open_workbook(filename='C:/Users/User/Desktop/数据底座_mapping_V1.1/'+i)
    sheet_count = len(workbook.sheet_names()) # 获取工作表数量
    for j in range(2,sheet_count):
        start_time = time.time()
        table = workbook.sheet_by_index(j) # 获取第j个工作表
        row = table.nrows # 获取行数
        col = table.ncols # 获取列数
        Cn_name = table.cell_value(rowx=2, colx=1) # 获取中文表名
        En_name = table.cell_value(rowx=1, colx=1) # 获取英文表名
        for r in range(9,row): # 从第9行开始遍历行
            for c in range(1,col): # 遍历列
                value = table.cell_value(rowx=r, colx=c) # 获取单元格的值
                ts.write(start_r,c+1,value,fieldNameStyle) # 将值写入表格
            ts.write(start_r,0,Cn_name,tableNameStyle) # 将中文表名写入第一列
            ts.write(start_r,1,En_name,tableNameStyle) # 将英文表名写入第二列
            start_r += 1
        delta_time = time.time()-start_time
        print('%s已加载完成'%table.name,'用时%s秒'%round(delta_time,4))
wtbook.save('C:/Users/User/Desktop/数据底座模型映射/数据底座模型映射文档-结果.xlsx')


            


                
                
