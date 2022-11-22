import encodings

import openpyxl

import read


def write_xlsx_excel(url , sheet_name , two_dimensional_data):
    '''
    写入xlsx格式文件
    参数：
        url:文件路径
        sheet_name:表名
        two_dimensional_data：将要写入表格的数据（二维列表）
    '''
    # 创建工作簿对象
    workbook = openpyxl.Workbook()
    # 创建工作表对象
    sheet = workbook.active
    # 设置该工作表的名字
    sheet.title = sheet_name
    # 遍历表格的每一行
    for i in range(0 , len(two_dimensional_data)):
        # 遍历表格的每一列
        for j in range(0 , len(two_dimensional_data[i])):
            # 写入数据（注意openpyxl的行和列是从1开始的，和我们平时的认知是一样的）
            sheet.cell(row=i + 1 , column=j + 1 , value=str(two_dimensional_data[i][j]))
    # 保存到指定位置
    workbook.save(url)
    print("写入成功")


f = open('read.txt' , 'r')
date = []
for i in f:
    date.append(i)
write_xlsx_excel('fengpan.xlsx' , 'fengpan' , date)
