# -*-coding:utf-8-*-
import requests
import json
from selenium import webdriver
import time
import xlwt
import xlrd
from xlutils.copy import copy
# 开始的页码
start_page = 717
# 一次爬多少页
max_page = 1


try:
    workbook = xlrd.open_workbook('./yanbao.xls',formatting_info=True)
    workbook = copy(workbook)
    worksheet = workbook.get_sheet(0)
except:
    print('无yanbao的sheet，新增？')
    ask='n'
    ask=input()
    if ask=='y':
        workbook = xlwt.Workbook(encoding = 'utf-8')
        worksheet = workbook.add_sheet('yanbao')
    else:
        exit()
dr = webdriver.Chrome(executable_path='./chromedriver')
dr.get('http://data.eastmoney.com/report/stock.jshtml')
#stock_table > table > tbody > tr:nth-child(1) > td:nth-child(1)
def get_page():
    for x in range(1,51):
        line_css=dr.find_element_by_css_selector('#stock_table > table > tbody > tr:nth-child('+str(x)+') > td:nth-child(1)')
        line=line_css.text
        for y in range(1,16):
            # try:#stock_table > table > tbody > tr:nth-child(1) > td:nth-child(1)
                selector='#stock_table > table > tbody > tr:nth-child('+str(x)+') > td:nth-child('+str(y)+')'
                # print(selector)
                element = dr.find_element_by_css_selector(selector)
                worksheet.write(int(line),y,element.text)
            # except :
            #     print(x,y)
            #     exit()
def next_page():
    next=dr.find_element_by_link_text('下一页')
    next.click()
def startpage(start_page):
#gotopageindex#gotopageindex
    putin = dr.find_element_by_css_selector('#gotopageindex')
    putin.send_keys(start_page)
    go = dr.find_element_by_css_selector('#stock_table_pager > div.gotopage > form > input.btn')
    go.click()
startpage(start_page)
time.sleep(1)
for i in range(1,max_page+1):
    print("当前"+str(start_page+i-1))
    get_page()
    next_page()
    workbook.save('yanbao.xls')
    # time.sleep()
dr.quit()
