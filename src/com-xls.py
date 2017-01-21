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
from readFile import *
from writeFile import *


# 此方法是用来测试脚本使用
def test():
    print("开始测试")
    test_file_1 = "test/test1.xlsx"
    test_file_2 = "test/test2.xlsx"
    readFile = ReadFile()
    writeFile = WriteFile()
    test_result = writeFile.copy_excel(test_file_2)
    rd_bk1 = readFile.open_excel(test_file_1)
    rd_bk2 = readFile.open_excel(test_result)
    wt_bk2 = copy(rd_bk2)
    test_dict_1 = readFile.get_dict_from_xls_file(test_file_1)
    test_dict_2 = readFile.get_dict_from_xls_file(test_result)
    for k1 in test_dict_1.keys():
        if k1 in test_dict_2.keys():
            list_1 = test_dict_1.get(k1)
            list_2 = test_dict_2.get(k2)
            if len(set(list_1) & set(list_2)):
                print(list_1)
                print(list_2)
    print("测试结束")
    pass

def main():
    test()

if __name__=="__main__":
    main()
