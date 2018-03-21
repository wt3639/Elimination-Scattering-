import xlrd
import xlwt
import sys
import random
from xlutils.copy import copy

for t in range(1,len(sys.argv)):
    #output = open(r'C:\Users\wutao\Desktop\point', 'w+')
    print(sys.argv[t]+' start processing...')
    xlsfile = sys.argv[t] # 打开指定路径中的xls文件
    book = xlrd.open_workbook(xlsfile)#得到Excel文件的book对象，实例化对象
    changebook = copy(book)
    sheet0 = book.sheet_by_index(0) # 通过sheet索引获得sheet对象
    change_sheet = changebook.get_sheet(0)
    nrows = sheet0.nrows    # 获取行总数
    print("row",nrows)
    ncols = sheet0.ncols    #获取列总数
    print("col",ncols)
    row_1 = sheet0.row_values(0)     # 获得第1行的数据列表
    col_1 = sheet0.col_values(0)     # 获得第1列的数据列表
    leftest = 7
    for m in range(1,nrows):
        for n in range(1,ncols):
            if(float(sheet1.cell_value(m,n)) <0):
                 change_sheet.write(m,n,0)     
    for i in range(1,nrows):
        for j in range(1,ncols):
            #if(float(sheet0.cell_value(i,j)) < 0):
                 #change_sheet.write(i,j,0)
            if(j>leftest and j<31 and i==1):
                for l in range(0,(j-leftest)):
                    max_num = sheet0.cell_value(j-leftest+1,j)
                    min_num = 0
                    writeValue = min_num+(max_num-min_num)/(j-leftest+1)*(l+1)
                    if(writeValue < 0):
                        change_sheet.write(i+l,j,0)
                    else:
                        change_sheet.write(i+l,j,writeValue)
                    #output.write(str(i+l)+' '+str(j)+'\n')
            if(row_1[j]==col_1[i] and j>30):
                for k in range(0,20):
                    min_num = sheet0.cell_value(i-4,j)
                    max_num = sheet0.cell_value(i+17,j)
                    writeValue = min_num+(max_num-min_num)/21*(k+1)
                    if(writeValue < 0):
                        change_sheet.write(i+l,j,0)
                    else:
                        change_sheet.write(i-3+k,j,writeValue)
                    #output.write(str(i-3+k)+' '+str(j)+'\n')
            if(row_1[j]*2 == col_1[i] and i>1):
                for n in range(0,24):
                    max_num = sheet0.cell_value(i-6,j)
                    if(i+18+1>nrows):
                        min_num = 0
                    else:
                        min_num = sheet0.cell_value(i+18,j)
                    if(i-5+n<nrows):
                        writeValue = max_num-(max_num-min_num)/25*(n+1)
                        if(writeValue < 0):
                            change_sheet.write(i+l,j,0)
                        else:
                            change_sheet.write(i-5+n,j,writeValue)
                        #output.write(str(i-5+n)+' '+str(j)+'\n')
                if(i>=nrows-2):
                    for m in range(j+1,ncols):
                        if(i-5+m-j<=nrows):
                            for n in range(0,24):
                                max_num = sheet0.cell_value(i-6+m-j,m)
                                if(i+18+1+m-j>nrows):
                                    min_num = 0
                                else:
                                    min_num = sheet0.cell_value(i+18+m-j,m)
                                if(i-5+n+m-j<nrows):
                                    writeValue = max_num-(max_num-min_num)/25*(n+1)
                                    if(writeValue < 0):
                                        change_sheet.write(i+l,j,0)
                                    else:
                                        change_sheet.write(i-5+n+m-j,m,writeValue)
                                    #output.write(str(i-5+n+m-j)+' '+str(m)+'\n')
                  
    save_file = sys.argv[t].split('.xls')[0]+'(Elimination Scattering).xls'
    changebook.save(save_file)
    print(sys.argv[t]+' processed completely')
#output.close()
print('All files complete!')



