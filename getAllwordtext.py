#!/usr/bin/python
# -*- coding: UTF-8 -*-
from tkinter import ttk           #  Tkinter 
import tkinter.messagebox
import tkinter.filedialog
import os.path
import sys, os
count = 0

### ###
import docx
line = '------------------------------\n'
def grep(s = ''):
	grepKey = e_input.get('0.0','end')	
	searchdir = e_path.get()
	fout = open('allwordtext','a',encoding='utf-8')
	for (dirname, dirs, files) in os.walk(searchdir):
		for filename in files:
			thefile = os.path.join(dirname,filename)  
			if filename.endswith('docx') :
				doccontent=readDocx(thefile)
				fout.write(line)
				fout.write(filename)
				fout.write('\n')
				fout.write(doccontent)   
	fout.close()

				      
#			for gkey in textcon:
#				if gkey !='':
#					gkey = gkey.lower()
#					gkey = gkey.strip()
#					print(gkey)
#					for line in doccontent:
#						line = line.lower()
#						line = line.strip()   
#						print(line)
#						if line.find(grepKey) != -1:          
#							count = count + 1          
#							print(line)
#
    		
	
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
    	return 1 	

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
	
#windows
window = tkinter.Tk()            #
window.title('TinyGrep 2.0')   #
window.geometry('500x400')  #
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
e_input.pack()


#grep dir 
l_path = ttk.Label(f2,text="select grep directory:",foreground="red",background="blue" )
l_path.pack(fill='x')

#grep dir
e_path = ttk.Entry(f2,text="...",)
e_path.pack(side = "left")

#grep dir choose button
but2 = ttk.Button(f2,text='path',command=getPath)
but2.pack(side = "right")             


#result dir
l_path_r = ttk.Label(f3,text="select grep result directory:",foreground="red",background="blue" )
l_path_r.pack(fill='x')

#result dir
e_path_r= ttk.Entry(f3,text="...",)
e_path_r.pack(side = "left")

#result dir
but3 = ttk.Button(f3,text='result',command=getPath_r) 
but3.pack(side = "right")             

#do grep 
but1 = ttk.Button(f4, 
    text='grep',         # 
    command=grep)     # 
but1.pack()       


#
window.mainloop()

