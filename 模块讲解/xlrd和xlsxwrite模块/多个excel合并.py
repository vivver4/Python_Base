import xlrd
import xlsxwriter
from datetime import date

def getFilect(shnum):
    '''
    读取文件每一行的内容并添加进list
    :param shnum:
    '''
    table = fh.sheets()[shnum]
    num = table.nrows
    for row in range(num):
        rdata = table.row_values(row)
        '''
        xlrd读取时间会转化为时间戳，因为通过下面方法复原
        '''
        if type(rdata[0]) == float:
            d = xlrd.xldate_as_tuple(int(rdata[0]), datemode=0)
            rdata[0] = date(*d[:3]).strftime('%Y-%m-%d')
        datavalue.append(rdata)


if __name__ == '__main__':
    '''
    想要批量打开的话可以用os模块的listdir一次性获取testdata下所有文件名
    '''
    allxls = ['D:/Python_Base/模块讲解/xlrd和xlsxwrite模块/testdata/多个excel合并1.xlsx',
              'D:/Python_Base/模块讲解/xlrd和xlsxwrite模块/testdata/多个excel合并2.xlsx']
    datavalue = []
    for f1 in allxls:
        fh = xlrd.open_workbook(f1)
        x = fh.nsheets
        for shnum in range(x):
            getFilect(shnum)

    endfile = xlsxwriter.Workbook('D:/Python_Base/模块讲解/xlrd和xlsxwrite模块/testdata/目标.xlsx')
    ws = endfile.add_worksheet('Sheet3')
    '''
    一个一个单元格来填入数据
    '''
    for a in range(len(datavalue)):
        for b in range(len(datavalue[a])):
            c = datavalue[a][b]
            ws.write(a, b, c)
    endfile.close()
    print('打印完成')