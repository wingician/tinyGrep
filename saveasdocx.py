import sys
import pickle
import re
import  codecs
import string
import shutil
from win32com import client as wc
import docx
 
 
def doSaveAas():
    word = wc.Dispatch('Word.Application')
    doc = word.Documents.Open(u'E:\code\\xxxx.doc')        # 目标路径下的文件
    doc.SaveAs(u'E:\\code\\hhhhhhhh.docx', 12, False, "", True, "", False, False, False, False)  # 转化后路径下的文件    
    doc.Close()
    word.Quit()
 
doSaveAas()
