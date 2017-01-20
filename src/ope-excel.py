# -*- coding: utf-8 -*-
import xlrd
import xlwt
import xlutils
import sys
import os
import types
import Levenshtein
from xlwt import *    
from readFile import *
from writeFile import *

from xlrd import open_workbook  
import sys  
from xlutils.copy import copy

'''
常用方法
Workbook----
1.get_sheet:获得sheet页

Worksheet----
1.set_name:修改sheet页名称

xlutils.display----
1.cell_display:读取cell内容
'''

#打开一个xls文件，读取数据
def open_excel(f= 'file.xls'):
    try:
        data = xlrd.open_workbook(f,encoding_override='utf-8')
        return data
    except Exception as e:
        print(e)

#创建一个xls文件，保存数据
def create_new_excel(filename):
    if (os.path.exists(filename)):
        print("文件已经存在,无需创建")
    else:
        w = Workbook(encoding='utf-8')
        w.add_sheet("sheet1")
        w.save(filename)

#复制一个excel表格
def copy_excel(soufile, desfile):
    if (not os.path.exists(soufile) or (os.path.exists(desfile))):
        print("请确保文件是否存在")
        return
    rb_wb = open_excel(soufile)
    wt_wb = copy(rb_wb)
    wt_wb.save(desfile)

#给一个表格插入列,每列的内容相同
def insert_one_col(wtsheet, rdsheet, newcolnum, instr, headstr=None):
    for moverow in range(0, rdsheet.nrows):
        for movecol in range(rdsheet.ncols-1, newcolnum-1, -1):
            if movecol == newcolnum:
                wtsheet.write(moverow, movecol, instr)
            else:
                wtsheet.write(moverow, movecol+1, rdsheet.cell(moverow, movecol).value.encode('utf-8'))

    if not headstr == None:
        wtsheet.write(0, newcolnum, headstr)

#给一个表格插入行,每列的内容相同
def insert_one_row(wtsheet, rdsheet, newrownum, instr):
    for movecol in range(0, rdsheet.ncols):
        for moverow in range(rdsheet.nrows-1, newrownum-1, -1):
            if moverow == newrownum:
                wtsheet.write(moverow, movecol, instr)
            else:
                wtsheet.write(moverow+1, movecol, rdsheet.cell(moverow, movecol).value.encode('utf-8'))

#输出整个Excel文件的内容  
def print_workbook(rb):  
    #sheet
    for s in rb.sheets():  
        #row
        for r in range(s.nrows):  
            #col
            for c in s.row(r):
                print(c.value)

#把excel的内容存到一个文件中
def wt_to_file(rb, filename):
    f = open(filename, 'w')
    for s in rb.sheets():
        for r in range(s.nrows):
            for c in range(s.ncols):
                f.write(str(r)+" " + str(c)+":"+str(s.cell(r, c).value.encode('utf-8')).replace('\n', '\t'))
            f.write('\n')
    f.close()

#获得两个excel 相同的sheet列表
def get_same_sheet(rb1, rb2):
    same_list = [sh for sh in rb1._sheet_names if sh in rb2._sheet_names]
    print(same_list)
    return same_list

#获取list列表中的字串成员
def get_str_list_from_list(des_list):
    str_list = [s for s in des_list if isinstance(s, types.StringTypes)]
    return str_list

#把一个list转换为一个字串，去掉其中的换行符
def get_str_from_list(des_list):
    return ''.join(des_list).replace('\n', '')

#合并两个行中的差异信息
def merge_df_row(row1, row2):
    pass

#获得两个sheet相似的列
def get_same_row(sheet1, sheet2, simi_val):
    row1_value0 = get_str_list_from_list(sheet1.row_values(0))
    row2_value0 = get_str_list_from_list(sheet2.row_values(0))
    same_col = []
    for col1 in row1_value0:
        for col2 in row2_value0:
            if Levenshtein.ratio(col1, col2) > simi_val:
                same_col.append(col1)
                break
    print(same_col)
    return same_col
    
#获得两个list中相似度达到一定条件的元素
def get_same_item(list1, list2, simi_val):
    list1 = sorted(get_str_list_from_list(list1))
    list2 = sorted(get_str_list_from_list(list2))
    same_col_list = []
    print("list1:")
    print(list1)
    print("list2:")
    print(list2)
    for item1 in list1:
        for item2 in list2:
            if Levenshtein.ratio(item1, item2) > simi_val:
                same_col = [item1, item2]
                same_col_list.append(same_col)
                break
    print(same_col_list)
    return same_col_list

#打印diff结果报表  
def print_report(report):  
    for o in report:  
        if isinstance(o, list):  
            for i in o:  
                print("\t" + i)  
    else:  
        print (o)  

#根据行比较两个sheet页
def diff_sheet_by_row(sheet1, sheet2, simi_val):
    row1_value = None
    row2_value = None
    for r1 in range(sheet1.nrows):  
        row1_value = sheet1.row_values(r1)
        row1_value = get_str_from_list(row1_value)
        print(sheet1.name + " sheet1 " + " row " + str(r1) + " " + row1_value)
        for r2 in range(sheet2.nrows):
            row2_value = sheet2.row_values(r2)
            row2_value = get_str_from_list(row2_value)
            print(sheet2.name + " sheet2 " + " row " + str(r2) + " " + row2_value)
            #相似度函数判断
            if Levenshtein.ratio(row1_value, row2_value) < simi_val:
                print("相等")


'''
比较两个sheet的差异度，合并相似度达到预期的行
sheet1-参与比较的sheet1
sheet2-参与比较的sheet2
simi_val-相似度预期值
diff_style-比较方式 0-行比较 1-单元格比较
'''
def diff_sheet(sheet1, sheet2, simi_val, diff_style=0):  
    row1 = None  
    row2 = None  
    if diff_style == 0:
        diff_sheet_by_row(sheet1, sheet2, simi_val)
    elif diff_style ==1:
        diff_sheet_by_cell(sheet1, sheet2, simi_val)

#diff两行  
def diff_row(row1, row2):  
    nc1 = len(row1)  
    nc2 = len(row2)  
    nc = max(nc1, nc2)  
    report = []  
    for c in range(nc):  
      ce1 = None;  
    ce2 = None;  
    if c<nc1:  
        ce1 = row1[c]  
    if c<nc2:  
        ce2 = row2[c]  

    diff = 0; # 0:equal, 1: not equal, 2: row2 is more, 3: row2 is less  
    if ce1==None and ce2!=None:  
        diff = 2  
        report.append("+CELL[" + str(c+1) + ": " + ce2.value)  
    if ce1==None and ce2==None:  
        diff = 0  
    if ce1!=None and ce2==None:  
        diff = 3  
        report.append("-CELL[" + str(c+1) + ": " + ce1.value)  
    if ce1!=None and ce2!=None:  
        if ce1.value == ce2.value:  
            diff = 0  
        else:  
          diff = 1  
          report.append("#CELL[" + str(c+1) + "]1: " + ce1.value)  
          report.append("#CELL[" + str(c+1) + "]2: " + ce2.value)  

    return report  


'''if __name__=='__main__':  
  if len(sys.argv)<3:  
    exit()  

  file1 = sys.argv[1]  
  file2 = sys.argv[2]  

  wb1 = open_workbook(file1)  
  wb2 = open_workbook(file2)  

  #print_workbook(wb1)  
  #print_workbook(wb2)  

  #diff两个文件的第一个sheet  
  report = diff_sheet(wb1.sheet_by_index(0), wb2.sheet_by_index(0))  
  print file1 + "\n" + file2 + "\n#############################"  
  #打印diff结果  
  print_report(report)  
'''
#对比两个表格差异


''' 得到一个excel的sheet个数
    rb: 已经打开的excel对象
'''
def xl_sheet_num(rb):
    count = len(b.sheets()) #sheet数量
    return count


''' 获得一个excel所有的sheet名字
    rb: 已经打开的excel对象
'''
def xl_sheet_name(rb):
    count = len(rb.sheets())
    for sheet in rb.sheets():
        print(sheet.name)#sheet名称

'''获得表格中某个sheet某行的数据
    file：Excel文件路径
    colnameindex：行号
    by_index：sheet 号
'''
def excel_table_byindex(data,by_index=0,rowindex=0):
    #通过索引顺序获取一个表
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数
    print(nrows,ncols)
    if rowindex in range(1,nrows):
        #行列数据
        row = table.row_values(rowindex)
        print("row===",row)
        app = {}
        print("row_length==",len(row))
        return row
    else:
        return null
'''对比两个表格的差异
'''
def excel_table_compare(rb_hw,rb_hq):
    hw_sheet_num = xl_sheet_name(rb_hw)
    hq_sheet_num = xl_sheet_name(rb_hq)
    for i in range(hw_sheet_num):
        table = rb_hw.sheets()[i]
        nrows = table.nrows
        ncols = table.ncols
        for j in range(nrows):
            com_string(rb_hw,rb_hq,)
'''
    匹配关键字
    compare:需要对比的excel
    com_sheet：需要对比的sheet
    com_row_index:需要对比的行
    source:参考文件
    sour_sheet:参考sheet
    sour_row_index:参考文件行

'''
def com_string_row_col(compare,source,
        com_sheet=0,com_row_index=0,sour_sheet=0,
        sour_row_index=0):
    com_row = excel_table_byindex(compare,com_sheet,com_row_index)
    sour_row = excel_table_byindex(source,sour_sheet,sour_row_index)
    for i in range(2,len(com_row)-2):
        com_string = com_row[i];
        if (com_string==""):
            continue
        print("com_string==",com_string)
        if com_string in sour_row:
            return 1
    else:
        return 0
'''

'''
def com_string_row_sheet(rb_hw,rb_hq,hw_sheet,hw_row,hq_sheet):
    hq_table = data.sheets()[by_index]
    nrows = table.nrows
    for i in range(nrows):
        if (com_string_row_col(rb_hw,rb_hq,hw_sheet,hw_row,hq_sheet,i) ==1):
            return 1
    return 0

def main():
    ok_file = "/home/zsi1989u/zsl-github/zsl-excle/test/ok_ex.xlsx"
    no_file = "/home/zsi1989u/zsl-github/zsl-excle/test1/no_ex.xlsx"
    txt = "/home/zsi1989u/zsl-github/zsl-excle/test/test.txt"
    sys_result = "/home/zsi1989u/zsl-github/zsl-excle/test/opt_telephony_search_sys_result.txt"
    readFile = ReadFile()
    writeFile = WriteFile()
    writeFile.write_list_to_file(["aaa"],no_file)
    readFile.get_list_from_file(ok_file)

    #获得读实例
    # rd_ok = open_excel(ok_file)
    # rd_no = open_excel(no_file)
    #获得写实例
    # wt_no = copy(rd_no)
    #插入一列内容
    #insert_one_col(wt_no.get_sheet(0),rd_no.sheet_by_index(0), 0, "11111111", "zzzzzz")
    # insert_one_row(wt_no.get_sheet(0),rd_no.sheet_by_index(0), 0, "11111111")
    # wt_no.save(no_file)
    # same_list = get_same_sheet(rd_ok, rd_no)
    # for sh_name in same_list:
        # sh1 = rd_ok.sheet_by_name(sh_name)
        # sh2 = rd_no.sheet_by_name(sh_name)
        #diff_sheet(sh1, sh2, 0.5, 0)
        # get_same_item(sh1.row_values(0), sh2.row_values(0), 0.1)

    #给一个表格插入五个默认列
    #现在有两种方案1.表格中的内容跟另外的一整个表格进行匹配，不按页进行区分，这个适合查找所有的评估记录
    #根据sheet页进行对比，这个适合查找某个小Ｔ责任人


if __name__=="__main__":
    main()
