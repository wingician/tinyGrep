#!/usr/bin/python
# -*- coding: UTF-8 -*-
from tkinter import ttk           #  Tkinter 
import tkinter.messagebox
import tkinter.filedialog

import os.path
import sys, os
import xlwt


class ExcelWrite(object):
    def __init__(self):
        self.excel = xlwt.Workbook()  # 创建一个工作簿
        self.sheet = self.excel.add_sheet('Sheet1')  # 创建一个工作表
    
    # 写入单个值
    def write_value(self, cell, value):
        '''
            - cell: 传入一个单元格坐标参数，例如：cell=(0,0),表示修改第一行第一列
        '''
        self.sheet.write(*cell, value)
        # （覆盖写入）要先用remove(),移动到指定路径，不然第二次在同一个路径保存会报错
        #os.remove(excel_path)
        
        
    # 写入多个值
    def write_values(self, cells, values):
        '''
            - cells: 传入一个单元格坐标参数的list，
            - values: 传入一个修改值的list，
            例如：cells = [(0, 0), (0, 1)],values = ('a', 'b')
            表示将列表第一行第一列和第一行第二列，分别修改为 a 和 b
        '''
        # 判断坐标参数和写入值的数量是否相等
        if len(cells) == len(values):
            for i in range(len(values)):
                self.write_value(cells[i], values[i])
        else:
            print("传参错误,单元格：%i个,写入值：%i个" % (len(cells), len(values)))
            

### ###
import docx
def grep(s = ''):      	
    #get all word docs in a text files
    f_doc_tmp = open('doctext','w',encoding='utf-8')
    searchdir = e_path.get()
    for (dirname, dirs, files) in os.walk(searchdir):
    	for filename in files:
    		thefile = os.path.join(dirname,filename)
    		if filename.endswith('docx'):
    			doccontent=readDocx(thefile)
    			f_doc_tmp.write(filename)  
    			f_doc_tmp.write('\n')  
    			if doccontent != False:
    				f_doc_tmp.write(doccontent)
    				f_doc_tmp.write('\n')
    f_doc_tmp.close()
    f_doc_tmp = open('doctext','r',encoding='utf-8')
    
    f_out_path = e_path_r.get()
    excel_path = f_out_path + r'/tinybooks.xls'
    excel = xlwt.Workbook(encoding = 'utf-8') 
    sheet = excel.add_sheet('Sheet1')
    sheet.write(0,0, 'word文档名')
    sheet.write(0,1, '包含文件')
    count = 0
    for line in f_doc_tmp:
    	if line.find('.docx') != -1:
    		fname = line.strip()
    	start1 = '《'
    	end1 = '》'
    	s = line.find(start1)
    	while s!=-1:
    		e = line.find(end1, s)
    		sub_str = line[s:e + len(end1)]
    		count = count + 1
    		sheet.write(count,0, fname)
    		sheet.write(count,1, sub_str)
    		s = line.find(start1, e)
    f_doc_tmp.close()
    excel.save(excel_path)
    tkinter.messagebox.showinfo(message="检索完毕！")
    if(os.path.exists("doctext")):
    	os.remove("doctext")
			
	
def readDocx(docName):                 #doc
    fullText = []                       
    try:
    	doc = docx.Document(docName)
    	paras = doc.paragraphs
    	for p in paras:
            fullText.append(p.text)
    	return '\n'.join(fullText)
    except:	
    	print('error',docName)   
    	return False 	

def getPath(s = ''):
#
	file_path = tkinter.filedialog.askdirectory()	
	e_path.delete(0,"end")
	e_path.insert('insert', file_path)
	    
	    
def getPath_r(s = ''):
#
	file_path_r = tkinter.filedialog.askdirectory()	
	e_path_r.delete(0,"end")
	e_path_r.insert('insert', file_path_r)
	

def clearPath(s = ''):
#
	e_path.delete(0,"end")

def clearPath_r(s = ''):
#
	e_path_r.delete(0,"end")


def clearInput(s = ''):
#
	e_input.delete(1.0,"end")
	
	
#windows
window = tkinter.Tk()            #
window.title('查找word文档内书名 1.0')   #
window.geometry('520x210')  #
f1 = tkinter.Frame(window)
f1.pack(side = "top", fill = 'both')
f2 = tkinter.Frame(window)      
f2.pack(side = "top", fill = 'both')    
f3 = tkinter.Frame(window)      
f3.pack(side = "top", fill = 'both') 
f4 = tkinter.Frame(window)      
f4.pack(side = "top", fill = 'both') 
   
#grep dir 
l_path = ttk.Label(f2,text="搜索目录:",foreground="black",background="blue" )
l_path.pack(fill='x')

#grep dir
e_path = ttk.Entry(f2,text="choose grep directory",)
e_path.pack(side = "left")

#grep dir clear button
b_grep_clear = ttk.Button(f2,text='clear',command=clearPath)
b_grep_clear.pack(side = "right")             


#grep dir choose button
b_grep = ttk.Button(f2,text='搜索..',command=getPath)
b_grep.pack(side = "right")             


#result dir
l_path_r = ttk.Label(f3,text="保存到:",foreground="black",background="blue" )
l_path_r.pack(fill='x')

#result dir
e_path_r= ttk.Entry(f3,text="choose result directory",)
e_path_r.pack(side = "left")

#grep dir clear button
b_result_clear = ttk.Button(f3,text='clear',command=clearPath_r)
b_result_clear.pack(side = "right")   

#result dir
b_result = ttk.Button(f3,text='结果..',command=getPath_r) 
b_result.pack(side = "right")  


#do grep 
b_grep = ttk.Button(f4, 
    text='检索',         # 
    command=grep)     # 
b_grep.pack()       


#info1
l_info1 = ttk.Label(f1,text="本工具用于检索目录下所有word文件（扩展名是.docx）内包含的其他书名文件。\n查找的规则是通过识别书名号《、》，\n 搜索结果会显示会打开一个excel表格纪录。",foreground="blue",background="blue" )
l_info1.pack(fill='x')

#info2
l_info2 = ttk.Label(f1,text="保存excle表格的内容形式是，每一行显示 搜索的文件名，其中包含的一个书名。\n同一文件内有多数文件的话将分行显示。",foreground="blue",background="blue" )
l_info2.pack(fill='x')
#
window.mainloop()

