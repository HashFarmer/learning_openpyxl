from openpyxl import Workbook, load_workbook

### 关于工作簿

# 创建、保存一个空白工作簿
wb = Workbook() 
wb.save("wb_openpyxl_3.xlsx") #所有针对工作簿的操作在此之前
wb.close()

#加载指定的工作簿
wb = load_workbook('hello_world.xlsx')


### 关于工作表

#创建新的工作表
sheet_a = wb.create_sheet('123')
#获取所有的工作表名称
sheet_names = wb.get_sheet_names()
print(sheet_names)

#选择指定的工作表
# 方法1
ws = wb.get_sheet_by_name(sheet_names[0])
# 方法2
ws2 = wb2['test']
# 方法3
ws3 = wb2.get_sheet_by_name('test')



### 关于单元格

#选择单个单元格
ws['A1']
ws.cell(1,1)

#选择单元格区域, 指针游走方向，从左到右，从上到下
ws['A1:B2']
ws['A1':'B2']

#<generator object Worksheet._cells_by_row>
ws.iter_rows('A1:C2')
#<generator object Worksheet._cells_by_col>
ws.iter_cols('A1:C2')
