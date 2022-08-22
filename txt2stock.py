#! /usr/bin/env python3


import os
import sys
import xlwt
import re
import fnmatch
import string
#import pandas as pd
#import random
#from distutils.core import setup
#import py2exe


def sys_init ():
    global NUMBER,ROW,COLUMN,DIR,OFILE,wb,ws;
    #print (sys.argv,len(sys.argv))
    #exit()
    ROW=0
    
    if len(sys.argv) == 1 :
        print('你没有输入参数，根据提示进行参数设置')
        DIR = input('请输入需要分析的数据目录：') 
        NUMBER = input('请输入需要分析的数据字段,用空格隔开：') 
        OFILE = input('请输入需要保存的excel文件：') 
        #NUMBER = NUMBER.split("\"")
        NUMBER = NUMBER.split()
        NUMBER=sorted(list(NUMBER),key=float)
        print('根据输入，你的分析目录为 :' ,DIR)
        print('根据输入，你的分析数据为 :' ,NUMBER)
        print('根据输入，你的保存数据为 :' ,OFILE)
        #print(NUMBER)
    elif len(sys.argv) == 4 :
        DIR=sys.argv[1]
        NUMBER=sys.argv[2]
        OFILE=sys.argv[3]
        NUMBER = NUMBER.split()
        NUMBER=sorted(list(NUMBER),key=float)
        print('根据输入，你的分析目录为 :' ,DIR)
        print('根据输入，你的分析数据为 :' ,NUMBER)
        print('根据输入，你的保存数据为 :' ,OFILE)
       #print('DIR,OFILE,NUM,OFILE',DIR,OFILE,NUMBER,OFILE)
        #exit()
    else :
        print('输入参数不对')
        exit()
 
def wb_init ():

    global NUMBER,ROW,COLUMN,DIR,OFILE,wb,ws;
     
    #COLUMN=['日期'	'买交易量B1'	'卖交易量S1'	'......' 'Total''买交易量Bn'	'卖交易量Sn'	'BM1'	'SM1' '......'	'BMn'	'SMn']
    COLUMN=[]
    COLUMN.insert(0,'文件名')
    COLUMN.insert(2*len(NUMBER)+1 , 'B[>' + str(NUMBER[-1]) + ']')
    COLUMN.insert(2*len(NUMBER)+2 , 'S[>' + str(NUMBER[-1]) + ']')
    COLUMN.insert(2*len(NUMBER)+3 , 'Total') # ADD Total
    #print(COLUMN)
    #print('NUMBER=',NUMBER)

    wb = xlwt.Workbook(encoding='gbk',style_compression=0)
    ws = wb.add_sheet('Sheet1',cell_overwrite_ok=True)
    

    for id in range(0,len(NUMBER)) :
        # print('id=',id)
        COLUMN.insert(2*id+1,'B[' + str(NUMBER[id]) + ']')
        COLUMN.insert(2*id+2,'S[' + str(NUMBER[id]) + ']')
        COLUMN.insert(2*len(NUMBER)+2*id+3+1 , 'BM[' + str(NUMBER[id]) + ']')
        COLUMN.insert(2*len(NUMBER)+2*id+4+1 , 'SM[' + str(NUMBER[id]) + ']')


        
        # print(2*id+1,2*id+2,2*len(NUMBER)+2*id+3,2*len(NUMBER)+2*id+4,4*len(NUMBER)+3,4*len(NUMBER)+4)
        # print(COLUMN,len(COLUMN))

    COLUMN.insert(4*len(NUMBER)+3+1 , 'BM[>' + str(NUMBER[-1]) + ']')
    COLUMN.insert(4*len(NUMBER)+4+1 , 'SM[>' + str(NUMBER[-1]) + ']')

    for i in range(0,len(COLUMN)):
        ws.write(0,i,COLUMN[i]) 
        
    #print(COLUMN,len(COLUMN))


def get_files(path):
    fs = []
    path=os.path.abspath(path)
    for root, dirs, files in os.walk(path):
        for file in files:
            if fnmatch.fnmatch(file, '*.txt'):
                fs.append(os.path.join(root, file))
    return fs

def ParLine (lines):
    global NUMBER,ROW,COLUMN,DIR,OFILE,wb,ws;
    #print('NUMBER=ParLine',NUMBER)

    for line in lines:
        str = line.split()
        #print('len(str)=',len(str),str)
        if len(str) == 5 :
            price=float(str[1]) * 100
            t_idx=2*len(NUMBER)+3
            #print('t_idx=',t_idx)
            dealNu=str[2]
            #print('num,price,Nu=',NUMBER,price,dealNu)
            NUMBER.append(dealNu)
            NUMBER=sorted(list(NUMBER),key=float)
            priceIndex = NUMBER.index(dealNu)
            NUMBER.remove(dealNu)
            #print('num=',NUMBER,price,dealNu)
            #new_str=sorted(list(new_str.append(-1.5)),key=float)
            #print('priceIndex=',priceIndex,NUMBER)
            B_index  =   2 * priceIndex + 1
            S_index  =   2 * priceIndex + 2
            BM_index =   2 * priceIndex + 3 + 1 + 2 * len(NUMBER)
            SM_index =   2 * priceIndex + 4 + 1 + 2 * len(NUMBER)
            #print('bi,si,smi,smi,len(nub)=',B_index,S_index,BM_index,SM_index,len(NUMBER))

            if  re.search("B", line):
                #print('B',len(str),str[1],str[2],str[3])
                #print(str)
                B_data    =                 float(dealNu) + float(COLUMN[B_index])
                BM_data   =  float(price) * float(dealNu) + float(COLUMN[BM_index])
                #print('B_data,BM_data',B_data,BM_data)
                COLUMN[B_index]     = B_data
                COLUMN[BM_index] = BM_data
                ws.write(ROW,B_index,B_data) 
                ws.write(ROW,BM_index,BM_data) 
                #t_data    =  float(B_data) ;
                #COLUMN[t_idx] = t_data
                #print('t_idx,t_data',t_idx,t_data)
                #ws.write(ROW,t_idx,t_data)
            elif re.search("S",line):
                #print('S,S_index,col,len(col)',str,S_index,COLUMN[S_index],len(COLUMN))
                #print('dealNu,S_index,SM_index,col(index)',dealNu,S_index,SM_index,COLUMN[S_index])

                S_data       =                          float(dealNu) + float(COLUMN[S_index])
                SM_data   =  float(price) * float(dealNu) + float(COLUMN[SM_index])
                #print('S_index,SM_data',S_index,SM_data)
                COLUMN[S_index]     = S_data
                COLUMN[SM_index] = SM_data
                ws.write(ROW,S_index,S_data) 
                ws.write(ROW,SM_index,SM_data) 
                #t_data    =  float(S_data) +  float(COLUMN[t_idx])
                #COLUMN[t_idx] = t_data
                #ws.write(ROW,t_idx,t_data)
            else :
                pass
        else :
            pass
    
    
    

    
    
# 合并文件
def merge():
    global NUMBER,ROW,COLUMN,DIR,OFILE,wb,ws;
    files = get_files(DIR)
    #print('files =',files)
    ROW=1
    for i in files:
        print('开始分析文件 : ',i)
        fopen = open(i)
        fid = os.path.basename(i).split('.')[0]#带后缀的文件名
        ws.write(ROW,0,fid)
        #print('filename=',fid)
        t_idx = 2*len(NUMBER)+3;
        for i in range(1,len(COLUMN)):
            # if (i == t_idx ):
                # print('t_idx',t_idx)
                # #COLUMN[i] = 'SUM(B1:t_idx）'
                # data='=SUM'+'('+'B'+str(str(ROW+1))+':'+chr(t_idx+64)+str(str(ROW+1))+')'
                # print('data',data)
                # ws.write(ROW,i,data) 
            # else :
                
                #ws.write(ROW,i,0) 
            COLUMN[i]=0.0

        #print('number=aaa',NUMBER)

        ParLine(fopen.readlines())
        #data = 0.0
        #T_idx = int(t_idx) + 1
        #print('idx',t_idx,T_idx)
        for j in range(1,t_idx) :
            #print('j',j)
            #print(int(COLUMN[j]),COLUMN[t_idx])
            COLUMN[t_idx] = int(COLUMN[j]) + int(COLUMN[t_idx])
            #print('data',data)
        #print('data',COLUMN[t_idx]);    
        ws.write(ROW,t_idx,COLUMN[t_idx]) 
        ROW=ROW+1

    print('\n分析完成，生成文件为：',os.path.abspath(OFILE))  
    wb.save(OFILE)

 

#print('number=',number)

if __name__ == '__main__':
    sys_init()
    wb_init()
    merge()
