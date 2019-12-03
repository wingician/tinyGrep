#!/usr/bin/python
# -*- coding: UTF-8 -*-
from tkinter import ttk           #  Tkinter 
import tkinter.messagebox
import tkinter.filedialog

import os.path
import sys, os
findone = "find one"
### ###
import docx
def grep(s = ''):      
	count = 0
	grepKey = e_input.get('0.0','end')
	f_key_tmp = open('grepkey','w',encoding='utf-8')
	f_key_tmp.write(grepKey)
	f_key_tmp.close()
	f_key = open('grepkey','r',encoding='utf-8')
	searchdir = e_path.get()
	for gkey in f_key:
		if gkey !='':
			gkey = gkey.lower()
			gkey = gkey.strip()
			print(gkey)	
			for (dirname, dirs, files) in os.walk(searchdir):
				for filename in files:
					thefile = os.path.join(dirname,filename)  
					if filename.endswith('docx') :
						doccontent=readDocx(thefile)
						if doccontent != False:
							f_doc_tmp = open('doctext','w',encoding='utf-8')
							f_doc_tmp.write(doccontent)                        
							f_doc_tmp.close()                               
							f_doc = open('doctext','r',encoding='utf-8')	
							for line in f_doc:
								line = line.lower()
								line = line.strip()   
								print(line)
								if line.find(gkey) != -1:          
									count = count + 1          
									print(findone)
							f_doc.close()
	f_key.close()
	if(os.path.exists("grepkey")):
		os.remove("grepkey")         
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
window.title('TinyGrep 2.0')   #
window.geometry('720x400')  #
f1 = tkinter.Frame(window)
f1.pack(side = "top", fill = 'both')
f2 = tkinter.Frame(window)      
f2.pack(side = "top", fill = 'both')    
f3 = tkinter.Frame(window)      
f3.pack(side = "top", fill = 'both') 
f4 = tkinter.Frame(window)      
f4.pack(side = "top", fill = 'both') 

 
 
#input eara
l_input = ttk.Label(f1,text="Please input keword:",foreground="red",background="white" )
l_input.pack(fill='x')

# input text(grep key word
e_input = tkinter.Text(f1,height=15,highlightcolor='black', highlightthickness=1)
e_input.pack(side = "left")


#grep dir clear button
b_input_clear = ttk.Button(f1,text='clear keywords',command=clearInput)
b_input_clear.pack(side = "right")      

#grep dir 
l_path = ttk.Label(f2,text="select grep directory:",foreground="red",background="blue" )
l_path.pack(fill='x')

#grep dir
e_path = ttk.Entry(f2,text="choose grep directory",)
e_path.pack(side = "left")

#grep dir clear button
b_grep_clear = ttk.Button(f2,text='clear',command=clearPath)
b_grep_clear.pack(side = "right")             


#grep dir choose button
b_grep = ttk.Button(f2,text='path..',command=getPath)
b_grep.pack(side = "right")             


#result dir
l_path_r = ttk.Label(f3,text="select grep result directory:",foreground="red",background="blue" )
l_path_r.pack(fill='x')

#result dir
e_path_r= ttk.Entry(f3,text="choose result directory",)
e_path_r.pack(side = "left")

#grep dir clear button
b_result_clear = ttk.Button(f3,text='clear',command=clearPath_r)
b_result_clear.pack(side = "right")   

#result dir
b_result = ttk.Button(f3,text='result..',command=getPath_r) 
b_result.pack(side = "right")             

#do grep 
b_grep = ttk.Button(f4, 
    text='grep',         # 
    command=grep)     # 
b_grep.pack()       


#
window.mainloop()

