# coding=utf-8
#  运行命令：python  final_excel.py  test.xls  excelwrite.xls 0,1,3,4,5,12   6,7,9,4  0,6,4,3  1

import xlrd
import xlwt

import argparse

# Training settings
parser = argparse.ArgumentParser(description='PyTorch MNIST Example')
##parser.add_argument('--input', type=str, default='test.xls', metavar='N',
##                    help='input excel (default: test.xls)')
##parser.add_argument('--output', type=str, default='excelwrite.xls', metavar='N',
##                    help='out excel (default: excelwrite.xls)')

parser.add_argument('input', default='test.xls')
parser.add_argument('output', default='excelwrite.xls')
parser.add_argument('mylist0', type=list )
parser.add_argument('mylist1', type=list )
parser.add_argument('mylist2', type=list )
parser.add_argument('flag', default='1' )

args = parser.parse_args()

args_cnt =  0
##for i in args:
print(args.input)
print(args.output)
print(args.mylist0)
print(args.mylist1)
print(args.mylist2)

mylist0 =   ''.join(args.mylist0)
mylist1 =   ''.join(args.mylist1)
mylist2 =   ''.join(args.mylist2)

mylist0=mylist0.split(",")
mylist1=mylist1.split(",")
mylist2=mylist2.split(",")
flag =  args.flag

print(mylist0)
print(mylist1)
print(mylist2)
print(flag)

if(flag=='0'):
    mylist0=''

##def read_xlrd(excelFile):
##    data = xlrd.open_workbook(excelFile)
##    table = data.sheet_by_index(0)
##    for rowNum in range(table.nrows):
##        rowVale = table.row_values(rowNum)
##        for colNum in range(table.ncols):
##            if rowNum > 0 and colNum == 0:
##                print(int(rowVale[0]))
##            else:
##                print(rowVale[colNum])
##        print("---------------")
##    # if判断是将 id 进行格式化
##    # print("未格式化Id的数据：")
##    # print(table.cell(1, 0))
##    # 结果：number:1001.0

# 打开文件
##data = xlrd.open_workbook('./test.xls')
data = xlrd.open_workbook(args.input)

workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('Sheet1')


# 查看工作表
data.sheet_names()
print("sheets：" + str(data.sheet_names()))

# 通过文件名获得工作表,获取工作表1
table  = data.sheet_by_name('Sheet2')
table2 = data.sheet_by_name('Sheet3')

# 打印data.sheet_names()可发现，返回的值为一个列表，通过对列表索引操作获得工作表1
# table = data.sheet_by_index(0)

# 获取行数和列数
# 行数：table.nrows
# 列数：table.ncols
print("总行数：" + str(table.nrows))
print("总列数：" + str(table.ncols))

# 获取整行的值 和整列的值，返回的结果为数组
# 整行值：table.row_values(start,end)
# 整列值：table.col_values(start,end)
# 参数 start 为从第几个开始打印，
# end为打印到那个位置结束，默认为none
#print("整行值：" + str(table.row_values(0)))
#print("整列值：" + str(table.col_values(1)))

# 获取某个单元格的值，例如获取B3单元格值
start = table.cell(0,0).value
##print("第一行第二列的值：" + cel_B3)
##print(start)

end = table.cell(1,0).value
##print("第一行第二列的值：" + cel_B3)
##print(end)

##for i in list(table2.nrows):
##    print("整行值：" + str(table2.row_values(i)))
def align(str1, distance, alignment='right'):
    length = len(str1.encode('gbk'))
    space_to_fill = distance - length if distance > length else 0
    if alignment == 'left':
        str1 = str1 + ' '* space_to_fill
    elif alignment == 'right':
        str1 = ' ' * space_to_fill + str1
    elif alignment == 'center':
        str1 = ' ' * (distance // 2) + str1 + ' ' * (distance - distance // 2)
    str1=str(str1)
    return str1

fline = ''
for col in range(table2.ncols):
##   if (col<2):
##       fline1 = align(str(table2.cell(0, col).value),20)
##       ##print('tmp2',fline1)
##       #.ljust(20, ' ')
##       fline = fline + ''+ fline1
##   elif (col==2):
##       fline = fline
##   elif  (col==3):
##       fline1 = align(str(table2.cell(0, col).value),35)
##       ##print('col3_len',len(str(table2.cell(0, col).value)));
##       ##print('tmp2',fline1)
##       fline = fline + ''+ fline1  +  '   '
##   elif  (col<6):
##       fline1 = align(str(table2.cell(0, col).value),20)
##       ##print('tmp2',fline1)
##       fline = fline + ''+ fline1
##   elif (col==12):
##       fline1 = align(str(table2.cell(0, col).value),20)
##       ##print('tmp2',fline1)
##       fline = fline + ''+ fline1
    if (str(col) in mylist0):
        print('col',col)
        fline1 = align(str(table2.cell(0, col).value), 20)
        fline = fline + '' + fline1
##print('fline',fline)


for row in range(table2.nrows):
    ##print("整行值：" + str(table2.row_values(row)))
    ##print(str(table2.row_values(row)))
    if (row>0):
        tmp=str(start)+ '''\n'''
        tmp = tmp +  fline  + '''\n'''
        for col in range(table2.ncols):
            ##if (col ==0):
            ##    tmp2 = align(str(table2.cell(row,col).value),20)
            ##    ##print('tmp2', tmp2)
            ##    tmp = tmp + '' + tmp2
            ##elif (col == 1):
            ##    ctype = table2.cell(row,col).ctype
            ##    ##print('ctype',ctype)
            ##    if(ctype==1):
            ##        tmp2 = align(str(table2.cell(row, col).value),20)
            ##    else :
            ##        tmp2 = align(str(int(table2.cell(row,col).value)),20)
            ##    ##print('tmp2', tmp2)
            ##    tmp = tmp + ''+ tmp2
            ##elif (col==2):
            ##    tmp=tmp
            ##elif (col ==3):
            ##    tmp2 = align(str(table2.cell(row,col).value),35)
            ##    ##print('tmp2', tmp2)
            ##    col3_len = len(str(table2.cell(row, col).value))
            ##    ##print('col3_len', len(str(table2.cell(row, col).value)));
            ##    tmp = tmp + ''+ tmp2  +  '   '
            ##elif (col < 6):
            ##    tmp2 = align(str(table2.cell(row,col).value),20)
            ##    ##print('tmp2', tmp2)
            ##    if (col3_len>10):
            ##        tmp = tmp  + '' + tmp2
            ##    else:
            ##        tmp = tmp + '  ' + tmp2
            ##elif (col == 12):
            ##    tmp2 = align(str(table2.cell(row, col).value),20)
            ##    ##print('tmp2', tmp2)
            ##    tmp = tmp + '' + tmp2
            if (str(col) in mylist0):
                tmp2 = align(str(table2.cell(row, col).value),20)
                tmp = tmp + '' + tmp2

            ##if(col==6):
            ##    tmp0=str(table2.cell(row,col).value)
            ##if (col == 7):
            ##    tmp6 = str(table2.cell(row, col).value)
            ##if (col == 9):
            ##    tmp4 = str(table2.cell(row, col).value)
            ##if (col == 4):
            ##    tmp3 = str(table2.cell(row, col).value)
            ##    tmp3_2 = ''
            ##    ack=tmp3.find('三级')
            ##    if(ack>=0):
            ##        tmp3_2 = '1、关键能力评议模版及说明2020.zip'
            ##    ack = tmp3.find('四级')
            ##    if (ack >= 0):
            ##        tmp3_2 = '1、关键能力评议模版及说明2020.zip'
            ##    ##print('tmp3_2',tmp3_2)
            if (str(col) in mylist1):
                ##print('col',col)
                tmp_list1 = str(table2.cell(row, col).value)
                worksheet.write(row, int(mylist2[int(mylist1.index(str(col)))]), str(tmp_list1))
                ##print(col,mylist2[int(mylist1.index(str(col)))])

        tmp = tmp + '''\n''' + str(end)
        ##print('tmp',tmp)
        worksheet.write(row, 2, str(tmp.encode('utf-8').decode('utf-8')))
        ##worksheet.write(row, 0, str(tmp0))
        ##worksheet.write(row, 4, str(tmp4))
        ##worksheet.write(row, 6, str(tmp6))
        ##worksheet.write(row, 3, str(tmp3_2))
        ##worksheet.write(row, 1, str('【请查收】FY19任职资格申请等级计划'))


##workbook.save('excelwrite.xls')
workbook.save(args.output)

print("OK".ljust(10, '+'))



