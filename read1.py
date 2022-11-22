import openpyxl


def read_xlsx_excel(url , sheet_name):
    '''
    读取xlsx格式文件
    参数：
        url:文件路径
        sheet_name:表名
    返回：
        data:表格中的数据
    '''
    # 使用openpyxl加载指定路径的Excel文件并得到对应的workbook对象
    workbook = openpyxl.load_workbook(url)
    # 根据指定表名获取表格并得到对应的sheet对象
    sheet = workbook[sheet_name]
    # 定义列表存储表格数据
    data = []
    # 遍历表格的每一行
    for row in sheet.rows:
        # 定义表格存储每一行数据
        row.replace('None')
        da = []
        # 从每一行中遍历每一个单元格
        for cell in row:
            # 将行数据存储到da列表
            da.append(cell.value)
            #da.replace('None')
            ''.join(da)
            print(da)
        # 存储每一行数据
        data.append(da)
    # 返回数据
    return data


b = read_xlsx_excel('缝盘2组6月份.xlsx' , '缝盘2组')

