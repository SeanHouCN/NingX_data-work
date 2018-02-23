#!/usr/bin/env python
# -*- coding: gbk -*-
"""
Created on Mon Feb  5 16:48:18 2018
Changed on Fri Feb  22
@author: HOU Jinxiu
"""

import os,shutil
import re 
import pandas as pd 
'''     
               文件名称
                   cwd LF文件上上层文件夹的绝对路径
'''
def filename(cwd):
    FF = []
    for filename in os.listdir(cwd): 
        if os.path.isdir(os.path.join(cwd,filename)):
            FF.append(filename)
    return FF

'''     
               读取程序
'''
def read_lf(lf_dir):
#    print("Current Working Path is %s"%lf_dir)
    fopen=open(lf_dir)
    lines =fopen.readlines(encoding=gbk)
    fopen.close()
#    print("Reading over")
    data = []
    for line in lines:
        line = re.split(r'[\,\s\'\;]+',line)
        data.append(line)
    m_data = pd.DataFrame.from_records(data)
#    print("Converting df over")
    return m_data
'''     
               数据整合程序
                   cwd 
                   filename 为filename函数获取的list其中之一
'''
def select_data(cwd,filename,result_dir):
    '''     
                   L1文件提取发电机编号
    '''
    lf_dir1 =cwd+"/"+filename+"/LF.L1"
    L1 = read_lf(lf_dir1)
    L1.drop([2,4,5,6,7,9,10],axis=1,inplace=True)
    for i  in range(len(L1)):
        L1.iloc[i,0] =i+1
    L1.columns=['num','long_name','area','short_name']
#    new_L1 = L1[L1['short_name'].str.contains(r'.*宁夏.*')]
    L1 = L1.set_index(['num'])
    
    '''    
            对比L2、LP2中的线路，实现I、J的名称映射
    '''
    lf_dir2 = cwd+"/"+filename+"/LF.L2"
    L2 = read_lf(lf_dir2)
    L2.drop([0,5,6,7,8,9,10,11,12,13,14,15,17],axis=1,inplace=True)
    L2.columns=['Mark','I','J','No','Name']
    L2['I'] = L2['I'].map(lambda x: L1.loc[int(x),'short_name'])
    L2['J'] = L2['J'].map(lambda x: L1.loc[int(x),'short_name'])
    new_L2 = L2[L2['Name'].str.contains(r'.*宁夏.*')]
    
    
    lf_dir3 = cwd+"/"+filename+"/LF.LP2"
    LP2 = read_lf(lf_dir3)
    LP2.drop([0,5,7,8,9,10],axis=1,inplace=True)
    LP2.columns=['I','J','No','Pi','Pj']
    LP2['I'] =LP2['I'].map(lambda x: L1.loc[int(x),'short_name'])
    LP2['J'] =LP2['J'].map(lambda x: L1.loc[int(x),'short_name'])
    new_LP2 = LP2[LP2['I'].str.contains(r'.*宁夏.*')]
    new_LP2['No'] = new_LP2['No'].map(lambda x: L2.loc[L2[L2.No== x].index.tolist()[0],'Name'])
    
    '''    
                       查询LP3中的变压器
    '''
    lf_dir4 = cwd+"/"+filename+"/LF.LP3"
    LP3 = read_lf(lf_dir4)
    LP3.drop([0,5,7,8,9,10],axis=1,inplace=True)
    LP3.columns=['I','J','No','Pi','Pj']
    LP3['I'] =LP3['I'].map(lambda x: L1.loc[int(x),'short_name'])
    LP3['J'] =LP3['J'].map(lambda x: L1.loc[int(x),'short_name'])
    new_LP3 = LP3[LP3['I'].str.contains(r'.*宁夏.*')]
    
    '''    
                     查询LP5中的发电机数据
    '''
    lf_dir5 = cwd+"/"+filename+"/LF.LP5"
    LP5 = read_lf(lf_dir5)
    LP5.drop([0,4,5],axis=1,inplace=True)
    LP5.columns=['I','P','Q']
    LP5['I'] =LP5['I'].map(lambda x: L1.loc[int(x),'short_name'])
    new_LP5 = LP5[LP5['I'].str.contains(r'.*宁夏.*')]
    '''    
                       查询LP6中的负荷
    '''
    lf_dir6 = cwd+"/"+filename+"/LF.LP6"
    LP6 = read_lf(lf_dir6)
    LP6.drop([0,5],axis=1,inplace=True)
    LP6.columns=['I','No','PL','QL']
    LP6['I'] =LP6['I'].map(lambda x: L1.loc[int(x),'short_name'])
    new_LP6 = LP6[LP6['I'].str.contains(r'.*宁夏.*')] 
    new_LP6=new_LP6.apply(lambda x: pd.to_numeric(x, errors='ignore'))
    new_LP6_PL = new_LP6[new_LP6.loc[:,'PL']>0]
    new_LP6_PG = new_LP6[new_LP6.loc[:,'PL']<0]
    
    

    with pd.ExcelWriter(cwd+'/'+filename+'.xlsx',engine='xlsxwriter') as writer:
        new_L2.to_excel(writer,sheet_name='L2') 
        new_LP2.to_excel(writer,sheet_name='new_LP2') 
        new_LP3.to_excel(writer,sheet_name='new_LP3')
        new_LP5.to_excel(writer,sheet_name='new_LP5')
        new_LP6_PL.to_excel(writer,sheet_name='new_LP6_PL')
        new_LP6_PG.to_excel(writer,sheet_name='new_LP6_PG')
    
    shutil.move(cwd+'/'+filename+'.xlsx',result_dir)
    print("done")
    return 0 


if __name__ == '__main__':
    cwd = r"xxxx"  
    result_dir = r"xxxx"
    print(filename(cwd)[0])
    for fn in filename(cwd):
        select_data(cwd,fn,result_dir)
