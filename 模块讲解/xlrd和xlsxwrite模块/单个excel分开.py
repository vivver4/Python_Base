import pandas as pd
import xlrd
xlsfile='D:/Python_Base/模块讲解/xlrd和xlsxwrite模块/testdata//单个excel分开.xlsx'
book=xlrd.open_workbook(xlsfile)
path='C:/Users/shangya/Desktop'
'''
循环所有的sheets
'''
for sheet in book.sheets():
    print(sheet.name)
    target = path + '/' + str(sheet.name) + '.xlsx'
    '''
    将sheet单独写入单独的excel文件
    '''
    with pd.ExcelWriter(target) as f:
        '''
        这里要注意header和index的参数设置，比如to_excel的两个参数默认True，容易造成错误，具体看网页收藏
        '''
        df=pd.read_excel(xlsfile, sheet_name=sheet.name, header=None)
        df.to_excel(f, header=False, index=False)
