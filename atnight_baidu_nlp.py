from openpyxl import load_workbook
from openpyxl import workbook
from openpyxl.utils import FORMULAE
from aip import AipNlp
from pprint import pprint
from time import sleep
import json,time
import jsonpath
import os
import re


class BaiDuAI:

    def __init__(self):

        APP_ID = '20025949'
        API_KEY = 'A96xFckepcEf0xCb1biESDgC'
        SECRET_KEY = 'tdfk4h9iiQYWFReijqGyw91Xw0eGGIkc'
        self.client = AipNlp(APP_ID, API_KEY, SECRET_KEY)

        self.OPTIONS = {"type":12}  # 设定为“购物”

    def filter_emoji(self,desstr,restr=''):  # 去除表情编码
        try:  
            co = re.compile(u'[\U00010000-\U0010ffff]')  
        except re.error:  
            co = re.compile(u'[\uD800-\uDBFF][\uDC00-\uDFFF]')
        return co.sub(restr, desstr)

    def excel(self,path):
        workbook = load_workbook(filename = path)
        sheet1 = workbook.active
        print(sheet1)  # 确认工作表

        print(sheet1.dimensions)  # 工作表信息

        sheet1["B1"] = "Sentiment"
        sheet1["C1"] = "Abstract"
        sheet1["D1"] = "prop"
        sheet1["E1"] = "adj"

        x = 0
        y = 2

        for i in sheet1["A2:A10000000"]:  # 区域
            for j in i:
                if j.value == None:
                    break
                a = j.value
                if isinstance(a,str):
                    a = self.filter_emoji(a,restr='')
                c = self.run1(a)
                #解析json数据
                c_sentiment = jsonpath.jsonpath(c,'$..sentiment')
                c_abstract = jsonpath.jsonpath(c,'$..abstract')
                c_prop = jsonpath.jsonpath(c,'$..prop')
                c_adj = jsonpath.jsonpath(c,'$..adj')
                print(c_sentiment,c_abstract,c_prop,c_adj)
                #写入Excel
                sheet1["E{}".format(y)] = str(c_sentiment) if isinstance(c_sentiment,list) else "none"
                sheet1["F{}".format(y)] = str(c_abstract) if isinstance(c_abstract,list) else "none"
                sheet1["G{}".format(y)] = str(c_prop) if isinstance(c_prop,list) else "none"
                sheet1["H{}".format(y)] = str(c_adj) if isinstance(c_adj,list) else "none"
                sleep(0.5)  # QPS限制为2
                x += 1
                y += 1
                print("-" * 50,"%d" % x,"-" * 50)
        workbook.save(filename=path)

    def run1(self,content):
        '''
        通用版
        '''
        try:
            return self.client.commentTag(content,options=self.OPTIONS)  # 调用API结果
        except:
            return {'log_id': 888, 'items': [{'sentiment': '', 'abstract': '', 'prop':'', 'begin_pos': 0, 'end_pos': 0, 'adj': ''}]}

    def run2(self,content):
        '''
        定制版
        '''
        try:
            return self.client.commentTagCustom(content,options=self.OPTIONS)  # 调用API结果
        except:
            return {'log_id': 888, 'items': [{'sentiment': '', 'abstract': '', 'prop':'', 'begin_pos': 0, 'end_pos': 0, 'adj': ''}]}

if __name__ == '__main__':

    baidunlp = BaiDuAI()
    #baidunlp.excel(r"G:\testfiles\领夹麦克风评论汇总(2020-07-31 150611).xlsx")
    baidunlp.excel(path)

    print("-" * 50,"{0}".format(time.strftime('%Y-%m-%d %H:%M')),"-" * 50)