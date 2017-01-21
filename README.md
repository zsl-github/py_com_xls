# py-com-xls
此脚本重要是用来对两个表格进行操作，包括如下功能(后续还会根据需要拓展)

**合并两个单元格中重复的部分**

比如，一个表格中有些信息结果要移植到另一个表格中

## 模块介绍

### xlrd

此模块主要是用来对`xls`表格进行读取操作

**关键方法**

获得一个book实例，里面存放的是整个表格的信息

```python
open_workbook
```

----

获得一个表格table

```python
book.sheets()[0]
book.sheet_by_index(0)
sheet_by_name(u'Sheet1')
```
----

获取行数和列数

```python
nrows = table.nrows
ncols = table.ncols
```

循环行列表数据

```pthon
for i in range(nrows ):
    print table.row_values(i)
```
单元格

```python
cell_A1 = table.cell(0,0).value
cell_C4 = table.cell(2,3).value
```
使用行列索引

```python
cell_A1 = table.row(0)[0].value
cell_A2 = table.col(1)[0].value
```
简单的写入
```python
row = 0
col = 0
```
单元格值类型
* empty
* string
* number
* date
* boolean
* error

```python
# ctype = 1
# value = '单元格的值'
# xf = 0
table.put_cell(row, col, ctype, value, xf)

#单元格的值'
table.cell(0,0).value 
```


