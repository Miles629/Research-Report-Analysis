# -*-coding:utf-8-*-
from jqdatasdk import *
import xlwt
import xlrd
from xlutils.copy import copy
# import json
from datetime import datetime,timedelta

# 开始行和结束行
start = 1
end = 35824
auth('xxxxxx用户名', 'sdasdas密码')
# 查询是否连接成功
is_auth = is_auth()
print(is_auth)
def get_close(code,date):
    code = normalize_code(code)
    response = get_price(code, start_date=date, end_date=date, frequency='daily', fields=None, skip_paused=False, fq=None, count=None, panel=True, fill_paused=True)
    return response['close']
def get_preday(day):
    return (datetime.strptime(day,'%Y-%m-%d')+ timedelta(days = -1)).strftime('%Y-%m-%d')
def get_nextday(day):
    return (datetime.strptime(day,'%Y-%m-%d')+ timedelta(days = +1)).strftime('%Y-%m-%d')

try:
    workbook = xlrd.open_workbook('./gupiao.xls',formatting_info=True)
    workbook = copy(workbook)
    worksheet = workbook.get_sheet(0)
except:
    print('无gupiao的sheet，新增？')
    ask='n'
    ask=input()
    if ask=='y':
        workbook = xlwt.Workbook(encoding = 'utf-8')
        worksheet = workbook.add_sheet('gupiao')
    else:
        exit()
xlsrd = xlrd.open_workbook('./yanbao.xls',formatting_info=True)
sheetrd = xlsrd.sheet_by_index(0)
# 不存在的日期这么判断
# p = get_close('300760','2020-11-21')
# print(p.count())

save=0
for l in range(start,end):
    code = sheetrd.cell_value(l,2)
    pre_day = xlrd.xldate.xldate_as_datetime(sheetrd.cell(l,15).value, 0).strftime('%Y-%m-%d')
    to_day = xlrd.xldate.xldate_as_datetime(sheetrd.cell(l,16).value, 0).strftime('%Y-%m-%d')
    next_day = xlrd.xldate.xldate_as_datetime(sheetrd.cell(l,17).value, 0).strftime('%Y-%m-%d')   #直接转化为datetime对象
    print("当前：",l,"    代码：",code)
    try:
        pre_price = get_close(code,pre_day)
        print("pre-price get")
        to_price = get_close(code,to_day)
        print("to-price get")
        next_price = get_close(code,next_day)
        print("next-price get")
        worksheet.write(l,1,code)
        if pre_price.count()==0:
            worksheet.write(l,10,'notintrade')
            a = 1
            print("前天数据为空，将进行8次回溯")
            while a<=8:
                pre_day=get_preday(pre_day)
                pre_price=get_close(code,pre_day)
                print("回溯",a,"天")
                if pre_price.count()==0:
                    worksheet.write(l,10-a,'notintrade')
                    a=a+1
                else:
                    worksheet.write(l,10-a,pre_price[0])
                    break
        else:
            worksheet.write(l,10,pre_price[0])
        if to_price.count()==0:
            worksheet.write(l,11,'notintrade')
        else:
            worksheet.write(l,11,to_price[0])
        if next_price.count()==0:
            worksheet.write(l,12,'notintrade')
            a = 1
            print("前天数据为空，将进行8次后续访问")
            while a<=8:
                pre_day=get_preday(pre_day)
                pre_price=get_close(code,pre_day)
                print("后续",a,"天")
                if pre_price.count()==0:
                    worksheet.write(l,12+a,'notintrade')
                    a=a+1
                else:
                    worksheet.write(l,12+a,pre_price[0])
                    break
        else:
            worksheet.write(l,12,next_price[0])
    except:
        worksheet.write(l,0,'wrong!')
    if save==20:
        workbook.save('gupiao.xls')
        print("now",l)
        save=0
    else:
        save=save+1
print("done!")
workbook.save('gupiao.xls')

