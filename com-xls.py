# -*- coding: utf-8 -*-

from xlrd import *
from xlutils.copy import copy
from xlwt import *    
import Levenshtein
import os
import sys
import sys  
import types
import xlrd
import xlutils
import xlwt

# 文件头列表
def get_file_type(): 
    return { 
        "FFD8FF": "JPEG", 
        "47494638":"GIF",
        "49492A00":"TIFF",
        "424D":"Windows Bitmap",
        "41433130":"CAD",
        "38425053":"Adobe Photoshop",
        "7B5C727466":"Rich",
        "3C3F786D6C":"XML",
        "68746D6C3E":"HTML",
        "44656C69766572792D646174653A":"Email",
        "CFAD12FEC5FD746F":"Outlook Express",
        "2142444E":"Outlook",
        "D0CF11E0":"Word-Excel",
        "5374616E64617264204A":"MS Access",
        "89504E47": "PNG"} 

# 字节码转16进制字符串 
def bytes2hex(byte): 
    num = len(byte) 
    hexstr = u"" 
    for i in range(num): 
        #转化为十六进制
        t = u"%x" % byte[i] 
        if len(t) % 2: 
            #如果是一位的话，进行补零操作
            hexstr += u"0"
        hexstr += t 
    #转化为大写
    return hexstr.upper() 

# 获取文件类型 
def check_file_type(filename): 
    # 必需二制字读取 
    binfile = open(filename, 'rb') 
    tl = get_file_type() 
    ftype = 'unknown'
    numOfBytes = 100
    # 一个 "B"表示一个字节 
    hbytes = struct.unpack_from('B'*numOfBytes, binfile.read(200)) 
    f_hcode = bytes2hex(hbytes) 
    binfile.close() 
    for hcode in tl.keys(): 
        # 需要读多少字节 
        if f_hcode.startswith(hcode):
            ftype = tl[hcode] 
            return tl[hcode]
    return None

    '''
    判断一个list是否在一个list列表中
    '''
    def is_list_in_dict(self, dic, lis):
        pass

#打开一个xls文件，读取数据
def open_excel(f_excel):
    try:
        data = xlrd.open_workbook(f_excel,encoding_override='utf-8')
        return data
    except Exception as e:
        print(e)

#复制一个excel表格
def copy_excel(soufile, desfile=""):
    rb_wb = open_excel(soufile)
    wt_wb = copy(rb_wb)
    if desfile == "":
        desfile = os.path.splitext(soufile)[0] + "_" + "result" + os.path.splitext(soufile)[1]
    wt_wb.save(desfile)
    return desfile

'''
获得一个excel文件的所有内容
'''
def get_list_from_excel(filename):

    excel_list=list()
    try:
        rd_book = open_excel(filename)
        for s in rd_book.sheets():  
            for r in range(s.nrows):
                row_list = s.row_values(r)
                if not list_is_empty(row_list):
                    excel_list.append(s.row_values(r))
    except Exception as e:
        print(e)

    return excel_list

'''
判断一个列表是不是全部都是空元素
'''
def list_is_empty(ch_list):
    return (len(set(ch_list)) ==1 and list(set(ch_list))[0] in ["",'',' ',"\n"])

'''
判断两个列表的相似度
'''
def list_simi_value(list_1, list_2, value):
    set_result = set(list_1) & set(list_2)
    if len(set_result) >= value:
        return True
    else:
        return False

'''
判断一个list是否在另外一个list集合中有相似的元素
'''
def list_in_list_set(list_item, list_set, value):
    for temp in list_set:
        if list_simi_value(remove_list_empty_item(list_item), remove_list_empty_item(temp), value):
            return temp

    return None
    

'''
去掉list中的空元素
'''
def remove_list_empty_item(list_all):
    return [i for i in list_all if i != ""]

# 此方法是用来测试脚本使用
def test():
    print("开始测试")
    sour_list = get_list_from_excel("test1.xlsx")

    test_result = copy_excel("test2.xlsx")
    rd_bk2 = open_excel(test_result)
    wt_bk2 = copy(rd_bk2)

    for s_index in range(rd_bk2.nsheets):  
        rd_s = rd_bk2.sheet_by_index(s_index)
        wt_s = wt_bk2.get_sheet(s_index)
        for r in range(rd_s.nrows):
            row_value = rd_s.row_values(r)
            result = list_in_list_set(row_value, sour_list, 3)
            if result != None:
                if row_value[2] != result[3]:
                    wt_s.write(r,2,result[3])
    wt_bk2.save(test_result)
    print("测试结束")

def main():
    test()

if __name__=="__main__":
    main()
