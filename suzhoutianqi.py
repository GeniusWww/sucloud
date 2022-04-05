#Author:yasongwu
# -*- codeing = utf-8 -*-
# @Time :2022/3/24 22:10
# @File : suzhoutianqi.py
# @Software: PyCharm

import requests
from lxml import etree
import xlwt
import datetime
import time

class Weather(object):
    def __init__(self):
        self.url = 'http://222.92.146.58:10110/HuiMaiReporting/AQIOfHour.aspx'
        self.headers ={
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.74 Safari/537.36'
}
        self.proxies = {
                # 'https' : 'https://202.109.157.62:9000',
                'http': 'http://120.220.220.95:8085'
            }

    def get_data(self,url):
        response = requests.get(url,headers = self.headers,proxies = self.proxies)
        return response.content.decode()

    def parse_data(self,data):
        html = etree.HTML(data)
        el_list = html.xpath('//tr[@style="text-align:center;"]/td')
        datalist = []
        cl_list = []
        for el in el_list:
            data = el.xpath("./text()")[0]
            datalist.append(data)

        for i in range(0, len(datalist), 19):
            bl_list = datalist[i:i + 19]
            cl_list.append(bl_list)

        lenth = len(cl_list)
        return cl_list,lenth



    def sava_data(self,lenth,cl_list):
        book = xlwt.Workbook(encoding="utf-8", style_compression=0)
        worksheet = book.add_sheet('苏州市天气信息', cell_overwrite_ok=True)
        col = ("监测点", "时间", "SO2", "分指数", "NO2", "分指数", "PM10", "分指数", "CO", "分指数", "O3", "分指数", "PM2.5", "分指数", "AQI",
               "首要污染物", "空气质量级别", "空气质量类别", "其他")
        for i in range(0, 19):
            worksheet.write(0, i, col[i])
        for i in range(0, lenth):

            try:
                data2 = cl_list[i]
            except:
                print('第{0}条数据错误'.format(i))
            for j in range(0, 19):
                worksheet.write(i + 1, j, data2[j])
        nowtime = datetime.datetime.now()


        savapath = nowtime.strftime('%Y-%m-%d %H%M%S')+'.xls'
        book.save(savapath)

    def run(self):
        data = self.get_data(self.url)
        cl_list, lenth = self.parse_data(data)
        self.sava_data(lenth,cl_list)

if __name__ == '__main__':


        yuhuaxing = Weather()
        yuhuaxing.run()
        time.sleep(3600)








