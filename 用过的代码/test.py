# -*-coding:utf-8-*-
# from jqdatasdk import *
import xlwt
import xlrd
from xlutils.copy import copy
import json
from datetime import datetime,timedelta
# # ��ʼ�кͽ�����
start = 1
end = 5
# auth('15539577560', '577560')
# # ��ѯ�Ƿ����ӳɹ�
# is_auth = is_auth()
# print(is_auth)

# def get_close(code,date):
#     code = normalize_code(code)
#     response = get_price(code, start_date=date, end_date=date, frequency='daily', fields=None, skip_paused=False, fq=None, count=None, panel=True, fill_paused=True)
#     return response['close']

# try:
#     workbook = xlrd.open_workbook('./gupiao.xls',formatting_info=True)
#     workbook = copy(workbook)
#     worksheet = workbook.get_sheet(0)
# except:
#     print('��gupiao��sheet��������')
#     ask='n'
#     ask=input()
#     if ask=='y':
#         workbook = xlwt.Workbook(encoding = 'utf-8')
#         worksheet = workbook.add_sheet('gupiao')
#     else:
#         exit()


xlsrd = xlrd.open_workbook('./yanbao.xls',formatting_info=True)
sheetrd = xlsrd.sheet_by_index(0)
# �����ڵ�������ô�ж�
# p = get_close('300760','2020-11-21')
# print(p.count())

# save=0
for l in range(start,end):
    code = sheetrd.cell_value(l,2)
    pre_day = xlrd.xldate.xldate_as_datetime(sheetrd.cell(l,15).value, 0).strftime('%Y-%m-%d')
    to_day = xlrd.xldate.xldate_as_datetime(sheetrd.cell(l,16).value, 0).strftime('%Y-%m-%d')
    next_day = xlrd.xldate.xldate_as_datetime(sheetrd.cell(l,17).value, 0).strftime('%Y-%m-%d')   #ֱ��ת��Ϊdatetime����
    print((datetime.strptime(pre_day,'%Y-%m-%d')+ timedelta(days = -1)).strftime('%Y-%m-%d'))
    # print(pre_day-1)

    # print(next_day)
#     try:
#         pre_price = get_close(code,pre_day)
#         to_price = get_close(code,to_day)
#         next_price = get_close(code,next_day)
#         worksheet.write(l,1,code)
#         if pre_price.count()==0:
#             worksheet.write(l,2,'notintrade')
#         else:
#             worksheet.write(l,2,pre_price[0])
#         if to_price.count()==0:
#             worksheet.write(l,3,'notintrade')
#         else:
#             worksheet.write(l,3,to_price[0])
#         if next_price.count()==0:
#             worksheet.write(l,4,'notintrade')
#         else:
#             worksheet.write(l,4,next_price[0])
#     except:
#         worksheet.write(l,2,'wrong!')
#     if save==20:
#         workbook.save('gupiao.xls')
#         print("now",l)
#         save=0
#     else:
#         save=save+1
# workbook.save('gupiao.xls')

