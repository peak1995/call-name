import xlrd
import numpy as np
def main():
    workbook = xlrd.open_workbook('学生名单.xlsx')
    sheet1 = workbook.sheet_by_index(0)
    data = list(sheet1.col_values(0))
    print(data)
    data.pop(0)
    print(data)
    print(len(data))
if __name__ == '__main__':
	main()