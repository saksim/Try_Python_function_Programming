#!-*-encoding=utf-8-*-
import pandas as pd
import xlrd,os,re

# 构建读取路径以及所有文件的函数
def load_path(folder):
    # 返回该目录下的所有EXCEL文件
    fd = os.path.realpath(folder)
    all_data = []
    for f in os.listdir(fd):
        if f.endswith('xls') or f.endswith('xlsx'):
            all_data.append(os.path.join(fd,f))
    return all_data

def handle_excel(data):
    # 目标是将每个文件SHEET中的数据，返回为一个大列表，为下一步运用做准备
    data_all = []
    for x in data:
        book = xlrd.open_workbook(x)
        for sheet in book.sheets():
            # print sheet.name
            for i in range(sheet.nrows):
                data_all.append(sheet.row_values(i))
    return data_all

def dict_data(data):
    # 目标：将单元格内包含特定分隔符的数据解析出，并将键值对按照键与值对读入的数据遍历
    # 输出：返回解析成功的数据矩阵.
    # 得到每一行
    z = []
    for i in data:
        y ={}
        # 得到当前行的列，遍历得到每一行的每一个列的值
        for j in i:
            # 对值进行切割，如果没有被切割 则返回的数据为原数据，否则为分割的数据
            # 注：现在没完成对全角的；以及：进行处理，只完成了对半角的;:的处理
            a = re.split('[;]',j)
            for m in a:
                b = re.split('[:]',m)
                if len(b) > 1:
                    y[b[0]] = b[1]
                else:y[a[0]] = a[0]
        z.append(y)
    print pd.DataFrame(z)

if __name__ == "__main__":
    folder = 'G:\\program_data\\test_excel'
    x = handle_excel(load_path(folder))
    dict_data(x)
