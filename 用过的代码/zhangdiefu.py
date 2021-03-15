# -*-coding:utf-8-*-
import xlwt
import xlrd
from xlutils.copy import copy

start=1
# end =10
end=35824

readbook=xlrd.open_workbook('./gupiao.xls',formatting_info=True)
readsheet = readbook.sheet_by_index(0)
try:
    workbook = xlrd.open_workbook('./zhangdie.xls',formatting_info=True)
    workbook = copy(workbook)
    worksheet = workbook.get_sheet(0)
except:
    print('无sheet，新增？')
    ask='n'
    ask=input()
    if ask=='y':
        workbook = xlwt.Workbook(encoding = 'utf-8')
        worksheet = workbook.add_sheet('zhangdie')
    else:
        exit()


save=0
for line in range(start,end):
    # 初始化
    code = readsheet.cell(line,1).value
    worksheet.write(line,1,code)
    iswrong = readsheet.cell(line,0).value
    if iswrong == 'wrong!':
        continue
    pre_num='notintrade'
    mid_num='notintrade'
    next_num='notintrade'
# 获取数据
    a=0
    while pre_num=='notintrade':
        if a==8:
            break
        pre_num=readsheet.cell(line,10-a).value
        a=a+1
    a=0
    while next_num=='notintrade':
        if a==8:
            break
        next_num=readsheet.cell(line,12+a).value
        a=a+1
    mid_num=readsheet.cell(line,11).value
    # print(pre_num,mid_num,next_num)
    if pre_num=='notintrade' or next_num=='notintrade':
        worksheet.write(line,0,'wrong')
        continue
    worksheet.write(line,3,(next_num-pre_num)/pre_num)#计算前一天和后一天相比的涨跌幅作为研报影响
    if mid_num=='notintrade':
        continue
    else:
        worksheet.write(line,4,(mid_num-pre_num)/pre_num)#计算研报发出当天比前一天的影响
        worksheet.write(line,5,(next_num-mid_num)/mid_num)#计算研报发出后一天比当天影响
    if save==50:
        print("now,save:",line)
        workbook.save("zhangdie.xls")
print("done!")
workbook.save("zhangdie.xls")
