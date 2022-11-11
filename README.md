# learning_openpyxl
# 安装
python -m pip install openpyxl

#
https://zhuanlan.zhihu.com/p/342422919


1.openpyxl简介
openpyxl是用于读取/写入Excel 2010 xlsx/xlsm文件的Python库，也就是说openpyxl这个Python库不支持xls文件的读取和操作，如果在工作中遇到xls文件我们就不能使用这个库。官方说它的诞生是因为缺少可从Python本地读取/写入Office Open XML格式的库，为了方便大家就开发了这个库，这是非常棒的。

2.文件转换
上述提到openpyxl只能操作xlsx文件，当我们遇到xls文件的时候就需要进行转化，转换方式这里提供几种方案供大家参考：

方法一：手动打开xlsx文件，然后另存为xlsx类型的文件。

方法二：使用pywin32模块进行转换，示例代码如下：

import os
import win32com.client as win32
filename = r'C:\Users\XH\Desktop\1.xls'
Excelapp = win32.gencache.EnsureDispatch('Excel.Application')
workbook = Excelapp.Workbooks.Open(filename)
# 转xlsx时: FileFormat=51,
# 转xls时:  FileFormat=56,
workbook.SaveAs(filename.replace('xls', 'xlsx'), FileFormat=51)
workbook.Close()
Excelapp.Application.Quit()
# 删除源文件
# os.remove(filename)

# 如果想将xlsx的文件转换为xls的话，则可以使用以下的代码：
# workbook.SaveAs(filename.replace('xlsx', 'xls'), FileFormat=56)


方法三：使用pandas模块进行转换，代码如下：

import pandas as pd
filename = r'C:\Users\XH\Desktop\1.xls'
filename2 = r'C:\Users\XH\Desktop\1.xlsx'
read_res = pd.read_excel(filename)
read_res.to_excel(filename2, index=False)
方法三在很多情况下出现一定的错误，比如在很多时候因为源表格的问题会造成数据丢失类的错误。个人推荐使用第二种方法。

