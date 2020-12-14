from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from openpyxl import load_workbook
from openpyxl import workbook
from openpyxl.utils import FORMULAE
from openpyxl.drawing.image import Image
import pandas as pd
from lxml import etree
from time import sleep
from decimal import Decimal
import random,csv,requests,os,re,time
import pyautogui
from urllib import request
from baidu_textOCR import BaiduOCR


class JD():

    def __init__(self):

        '''
        第一次启动流程：
        --->win + R
        --->输入cmd，回车
        --->复制cd C:\Program Files\Google\Chrome\Application,回车
        --->复制chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenium\AutomationProfile"，回车
        --->关闭黑色命令窗口，第一次是需从刚才调出的浏览器用自己的账号登录淘宝京东，第二次有记录则检查登录状态即可。
        '''
        desired_capabilities = DesiredCapabilities.CHROME  # 修改页面加载策略
        desired_capabilities["pageLoadStrategy"] = "none"  # 注释这两行会导致最后输出结果的延迟，即等待页面加载完成再输出
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        chrome_driver = r"C:\Program Files\Google\Chrome\Application\chromedriver.exe" 
        self.driver = webdriver.Chrome(chrome_driver,options=options)
        script = '''
        Object.defineProperty(navigator, 'webdriver', {
        get: () => undefined
        })
        '''
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": script})
        self.ocr = BaiduOCR()

    def linksToCompletion(self,i):
        # 整合链接
        link = "https://item.jd.com/{}.html".format(i)
        return link

    def parse_id(self,url):
        #正则提取ID，消除所有标点符号
        partten = re.compile(r"\d+")
        partten = partten.findall(url)
        return "".join(partten)

    def requestUrlList(self,beginToList):
        # 解析beginToList并发送请求,经过id去重
        for i in list(set(beginToList.strip().split())):
            # 请求URL
            if str(i).find("html")==-1:
                self.driver.get(self.linksToCompletion(i))
            else:
                self.driver.get(i)
            sleep(4)
            pyautogui.scroll(50)
            try:
                # 抓取京东页面信息
                self.jingdong()
            except Exception as error1:
                print("商品可能下架或预售期：{}，或出现{}".format(i,error1))

    def jingdong(self):

        sku_content = []
        # 解析京东商品详情页信息
        item = {}
        # 店铺名称
        item["Shop"] = self.driver.find_element_by_xpath('//div[@id="crumb-wrap"]//div[@class="name"]').text
        # 判断是否自营
        try:
            item["proprietary"] = self.driver.find_element_by_xpath('//div[@class="name goodshop EDropdown"]').text
        except:item["proprietary"] = "非自营"
        # 商品ID
        item["ID"] = self.parse_id(self.driver.current_url)
        # 商品价格
        try:
            item["Price"] = self.driver.find_element_by_xpath('//span[@class="p-price"]').text
        except:
            item["Price"] = self.driver.find_element_by_xpath('//span[@class="p-price ys-price"]').text
        # 商品标题
        item["NameCommodity"] = self.driver.find_element_by_xpath('//div[@class="sku-name"]').text.strip().replace(" ","")
        # 商品类目(改为品牌)
        item["category"] = self.driver.find_element_by_xpath('//ul[@id="parameter-brand"]').text
        try:
            # 商品卖点
            item["keySellingPoint"] = self.driver.find_element_by_xpath('//div[@id="p-ad"]').text
        except:item["keySellingPoint"] = ""
        # 产品参数
        item["parameter"] = self.driver.find_element_by_xpath('//ul[@class="parameter2 p-parameter-list"]').text.strip().replace("\n","；")

        # 月销量(改为链接累计评价数)
        try:
            item["Monthly_sales"] = self.driver.find_element_by_xpath('//div[@id="comment-count"]').text
        except:item["Monthly_sales"] = self.driver.find_element_by_xpath('//div[@class="activity-message"]').text
        # 当前商品评论数
        sleep(1)
        self.driver.find_element_by_xpath('//div[@class="tab-main large"]/ul//li[5]').click()
        sleep(1.5)
        try:
            self.driver.find_element_by_xpath('//input[@id="comm-curr-sku"]').click()
            sleep(0.5)
            item["Cumulative_comments"] = self.driver.find_element_by_xpath('//ul[@class="filter-list"]').text
        except:
            item["Cumulative_comments"] = "跳过"
        # 创建文件夹
        path_new = self.create_file(path,item)
        # 主图,下载及命名
        mainFigureCounts = 1
        mainFigure = self.driver.find_elements_by_xpath('//div[@id="spec-list"]//li')
        for i in mainFigure:
            mainPhotoLink = i.find_element_by_xpath('.//img').get_attribute("src").replace('n5','n1')
            item["mainFigure{}".format(mainFigureCounts)] = mainPhotoLink.split('/')[-1]
            request.urlretrieve(mainPhotoLink,os.path.join(path_new,mainPhotoLink.split('/')[-1]))
            mainFigureCounts+=1
        print(item)
    #        # 这里调用百度ocr接口做主图ocr
    #        try:
    #            item["mainFigure{}_ocr".format(mainFigureCounts)] = self.ocr.picture_Path(os.path.join(path_new,mainPhotoLink.split('/')[-1])) 
    #            sleep(0.5)
    #        except:
    #            print("OCR错误")
    #            item["mainFigure{}_ocr".format(mainFigureCounts)] =""
    #        mainFigureCounts+=1
    #    sku_content.append(item)
    #    # SKU列表解析
    #    sku_list = self.driver.find_elements_by_xpath('//ul[@class="tm-clear J_TSaleProp tb-img     "]//li')
    #    for i in sku_list:
    #        item_sku = {}
    #        # SKU列表点击
    #        try:
    #            element=i.find_element_by_xpath('./a')
    #            self.driver.execute_script("arguments[0].click();",element)
    #            sleep(0.5)
    #        except Exception as result:
    #            print(result)
    #            return sku_content
    #        # SKU不为空才加入item
    #        if i.find_element_by_xpath('./a').text !="":
    #            # SKU名称
    #            item_sku["SKU"] = i.find_element_by_xpath('./a').text
    #            # SKU价格
    #            item_sku["Price"] = self.driver.find_elements_by_xpath('//span[@class="tm-price"]')[-1].text
    #        sku_content.append(item_sku)
    #    #print(sku_content)
    #    return path_new,sku_content

    def create_file(self,path,item):
    # 判断文件夹是否已经存在，若存在则跳过，不存在则创建
        inspection_path = os.path.join(path,"{0}_ID{1}_{2}".format(time.strftime('%Y-%m-%d'),item["ID"],item["Shop"]))
        if not os.path.exists(inspection_path):
            os.makedirs(os.path.join(path,"{0}_ID{1}_{2}".format(time.strftime('%Y-%m-%d'),item["ID"],item["Shop"])))
        return inspection_path

    #def depositedInExcel(self,path,sku_content=None):
    #    wb = workbook.Workbook()
    #    sheet1 = wb.active
    #    cellCount = 2
    #    sheet1["A1"] = "商品标题"
    #    sheet1["A2"] = sku_content[0]["NameCommodity"]
    #    sheet1["B1"] = "重要卖点"
    #    sheet1["B2"] = sku_content[0]["keySellingPoint"]
    #    sheet1["C1"] = "商品类目"
    #    sheet1["C2"] = sku_content[0]["category"]
    #    sheet1["D1"] = "店铺名称"
    #    sheet1["D2"] = sku_content[0]["Shop"]
    #    sheet1["G1"] = "商品ID"
    #    sheet1["G2"] = sku_content[0]["ID"]
    #    sheet1["H1"] = "月销量"
    #    sheet1["H2"] = sku_content[0]["Monthly_sales"]
    #    sheet1["I1"] = "累计评价"
    #    sheet1["I2"] = sku_content[0]["Cumulative_comments"]
    #    sheet1["J1"] = "商品参数"
    #    sheet1["J2"] = sku_content[0]["parameter"]
    #    # 这个图有的还真的没有
    #    try:
    #        sheet1["K1"] = "主图1"
    #        sheet1["K2"] = sku_content[0]["mainFigure1"]
    #        sheet1["L1"] = "主图1_OCR"
    #        sheet1["L2"] = sku_content[0]["mainFigure1_ocr"]
    #        sheet1["M1"] = "主图2"
    #        sheet1["M2"] = sku_content[0]["mainFigure2"]
    #        sheet1["N1"] = "主图2_OCR"
    #        sheet1["N2"] = sku_content[0]["mainFigure2_ocr"]
    #        sheet1["O1"] = "主图3"
    #        sheet1["O2"] = sku_content[0]["mainFigure3"]
    #        sheet1["P1"] = "主图3_OCR"
    #        sheet1["P2"] = sku_content[0]["mainFigure3_ocr"]
    #        sheet1["Q1"] = "主图4"
    #        sheet1["Q2"] = sku_content[0]["mainFigure4"]
    #        sheet1["R1"] = "主图4_OCR"
    #        sheet1["R2"] = sku_content[0]["mainFigure4_ocr"]
    #        sheet1["S1"] = "主图5"
    #        sheet1["S2"] = sku_content[0]["mainFigure5"]
    #        sheet1["T1"] = "主图5_OCR"
    #        sheet1["T2"] = sku_content[0]["mainFigure5_ocr"]
    #    except:pass
    #    # SKU需要单独处理
    #    sheet1["E1"] = "SKU名称"
    #    sheet1["F1"] = "SKU价格"
    #    for i in sku_content[1:]:
    #        # 解决SKU列表不等于显示SKU列表的问题
    #        try:
    #            sheet1["E{}".format(cellCount)] = i["SKU"]
    #            sheet1["F{}".format(cellCount)] = i["Price"]
    #            cellCount+=1
    #        except:pass

    #    path = os.path.join(path,"SKU_{0}_ID{1}_{2}.xlsx".format(time.strftime('%Y-%m-%d'),sku_content[0]["ID"],sku_content[0]["Shop"]))
    #    # 完成一个表
    #    print(sku_content[0]["Shop"],sku_content[0]["ID"])
    #    wb.save(filename=path)


if __name__=="__main__":

    jd = JD()
    # 设置本地存储文件夹路径
    path = r"D:\Testfiles\cheyongxiche"
    # 在此输入商品ID列表--->Ctrl+s 保存--->shift+alt+F5 启动
    beginToList = '''
    2672959
    100014740254
    45835830963
    10021812936822
        '''
    jd.requestUrlList(beginToList)
    print("-" * 50,"{0}".format(time.strftime('%Y-%m-%d %H:%M')),"-" * 50)