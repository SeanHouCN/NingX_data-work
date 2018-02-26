#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Created on Stu Feb  24 23:04:18 2018
changed on Mon Feb  26 
@author: HOU
"""

import os,shutil
import re 
import pandas as pd 
import numpy as np
'''     
               提取文件名称
                   输入：cwd: LF文件上上层文件夹的绝对路径
                   输出：FF 文件名称
'''
def filename(cwd):
    FF = []
    for filename in os.listdir(cwd): 
        if os.path.isfile(os.path.join(cwd,filename)):
            FF.append(filename)
    return FF
'''     
               装换成日期格式
                   输入：X 传入的文件名参数
                   输出：日期和时间（2017-08-01  00-00-00）
'''
def f_split(X):
    X = X.strip('.xlsx').replace('_','-')
    return re.split('[T]',X)

'''     
                 线路差异对比输出
                   输入：df_PG_LP2: 本时刻的线路数据库
                        
                   输出：若有差异，则输出线路差异名称前增加是“+”（“-”）
                        dataframe格式数据[num（±），
                                        name（[减少的线路名称][增加的线路名称]）
                                        time 2017_08_01_00_15_00]          
'''
def compare_line(cwd,FF,df_PG_LP2,index):
    num = [] 
    time = []
    name = []
    if index != 0:
        path_xlsx = os.path.join(cwd,FF[index-1])
        df_PG_LP2_0 = pd.read_excel(path_xlsx,sheetname = "new_LP2")
        if df_PG_LP2.shape[0] != df_PG_LP2_0.shape[0]:            
            num_0 = df_PG_LP2.shape[0]-df_PG_LP2_0.shape[0]
            num.append(num_0)
            time.append(FF[index].strip('.xlsx').replace('T','_'))
#            if num_0 < 0:#线路减少了
            name1 = list(df_PG_LP2_0.No[~df_PG_LP2_0.No.isin(df_PG_LP2.No)])
#            else:      #线路增加了
            name2 = list(df_PG_LP2.No[~df_PG_LP2.No.isin(df_PG_LP2_0.No)])
            name.append(str(name1)+str(name2))
    c={"time" : time,"num" : num, "name" :name}
    df_compare_line = pd.DataFrame(c)
#    print(df_compare_line)
    return df_compare_line


'''     
                数据读取程序
                   输入：cwd: LF文件上上层文件夹的绝对路径
                        FF工作路径下的文件名list
                   输出：df_compare_line 线路的变换情况
                        df_all  电网整体概况
                        [PG_sum : 发电机总出力； G_size ：发电机个数；
                        line_size： 线路总条数； Pline_max,最大功率线路名称；
                        some_importent_lines：断面功率；sum_Load ：总负荷]6列
                        
'''
def read_excel(cwd,FF):
    df_all = pd.DataFrame(np.zeros((len(FF),4)),columns=['PG_sum','G_size','line_size','sum_Load'])
    df_compare_line =pd.DataFrame(columns=['name','num','time'])
    index = 0
    for ff in FF:
        path_xlsx = os.path.join(cwd,ff)
        print('这是第%s个，文件名是：%s'%(index+1,ff))
        ### 1.PG_sum : 发电机总出力
        df_PG_LP5 = pd.read_excel(path_xlsx,sheetname = "new_LP5")
        df_PG_LP6 = pd.read_excel(path_xlsx,sheetname = "new_LP6_PG")
        df_all['PG_sum'][index] = df_PG_LP5['P'].sum() - df_PG_LP6['PL'].sum()
        
        ### 2.PG_sum : 发电机总出力
        df_all['G_size'][index] = df_PG_LP5.shape[0]+df_PG_LP6.shape[0]
                
        ### 3.line_size： 线路总条数 
        df_PG_LP2 = pd.read_excel(path_xlsx,sheetname = "new_LP2")
        df_all['line_size'][index] = df_PG_LP2.shape[0]
        df_c = compare_line(cwd,FF,df_PG_LP2,index)
        if index>0:            
            if not df_c.empty:
                df_compare_line = pd.concat([df_c,df_compare_line])
                print(df_compare_line)
        ### 4.sum_Load 总负荷 
        df_new_LP6_PL = pd.read_excel(path_xlsx,sheetname = "new_LP6_PL")
        df_all['sum_Load'][index] = df_new_LP6_PL.shape[0]
        
        index += 1
        
    return df_all,df_compare_line








def main():
    cwd = r"C:\Users\HOU\Desktop\test_result"
    FF = filename(cwd)
    df_data = pd.DataFrame(list(map(f_split,FF)))
    df_all,df_compare_line = read_excel(cwd,FF)
    data = pd.concat([df_data,df_all],axis=1)

######test compare_line功能
#    path_xlsx = os.path.join(cwd,FF[114])
#    df_PG_LP2 = pd.read_excel(path_xlsx,sheetname = "new_LP2")
#    index = 114
#    df_compare_line = compare_line(cwd,FF,df_PG_LP2,index)
    
       
    
#    with pd.ExcelWriter(r'C:\Users\HOU\Desktop\统计数据.xlsx',engine='xlsxwriter') as writer:
#        pd.concat([df_data,df_all],axis =1).to_excel(writer,sheet_name='统计数据') 
    return data,df_compare_line

if __name__ == '__main__':
     data, df_compare_line = main()
#     data_08 = data[data.iloc[:,0].str.contains(r'^2017-08.*')]
