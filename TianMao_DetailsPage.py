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


class TianMao:

    def __init__(self):

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
        link = "https://item.taobao.com/item.htm?id="+ i
        return link

    def eliminateIllegalCharacter(self,name):
        # 解决非法字符问题
        rstr = r"[\/\\\:\*\?\"\<\>\|]"  # '/ \ : * ? " < > |'
        name = re.sub(rstr,"_",name)  # 替换为下划线
        return name


    def parse_id(self,url):
        #正则提取ID，消除所有标点符号
        partten = re.compile(r"id=(\d+)")
        partten = partten.findall(url)
        return "".join(partten)

    def create_file(self,path,item):
    # 判断文件夹是否已经存在，若存在则跳过，不存在则创建
        inspection_path = os.path.join(path,"{0}_ID{1}_{2}".format(time.strftime('%Y-%m-%d'),item["ID"],item["Shop"]))
        if not os.path.exists(inspection_path):
            os.makedirs(os.path.join(path,"{0}_ID{1}_{2}".format(time.strftime('%Y-%m-%d'),item["ID"],item["Shop"])))
        return inspection_path

    def requestUrlList(self,beginToList):
        # 解析beginToList并发送请求,经过id去重
        for i in list(set(beginToList.strip().split())):
            # 请求URL
            if str(i).find("com")==-1:
                self.driver.get(self.linksToCompletion(i))
            else:
                self.driver.get(i)
            sleep(random.randint(4,5))
            pyautogui.scroll(30)
            try:
                try:
                    # 抓取天猫页面信息
                    path_new,sku_content = self.tianmao(path)
                except Exception as error1:
                    # 抓取淘宝页面信息
                    print("error1：淘宝商品",error1)
                    path_new,sku_content = self.taobao(path)
                # 存入Excel
                self.depositedInExcel(ocrSwitch=ocrSwitch,path=path_new,sku_content=sku_content)
            except Exception as error2:
                print("商品可能下架或预售期：{}，或出现{}".format(i,error2))
            # 存入Excel，这位置调试
            #self.depositedInExcel(path=path_new,sku_content=sku_content)

    def tianmao(self,path):

        sku_content = []
        # 解析天猫商品详情页信息
        item = {}
        # 店铺名称
        try:
            item["Shop"] = self.driver.find_element_by_xpath('//a[@class="slogo-shopname"]').text
        except:item["Shop"] = ""
        # 商品ID
        item["ID"] = self.parse_id(self.driver.current_url)  
        try:
            # 商品卖点
            item["keySellingPoint"] = self.driver.find_element_by_xpath('//div[@class="tb-detail-hd"]/p').text
        except:item["keySellingPoint"] = ""
        # 商品标题
        item["NameCommodity"] = self.driver.find_element_by_xpath('//div[@class="tb-detail-hd"]/h1').text
        # 商品类目
        item["category"] = self.driver.find_element_by_xpath('//span[@class="yk-font-red-color yk-other-cat"]').text
        # 产品参数
        item["parameter"] = self.driver.find_element_by_xpath('//ul[@id="J_AttrUL"]').text.strip().replace("\n","；")
        # 月销量
        item["Monthly_sales"] = self.driver.find_element_by_xpath('//ul[@class="tm-ind-panel"]//li[@class="tm-ind-item tm-ind-sellCount"]').text
        # 评论数
        try:
            item["Cumulative_comments"] = self.driver.find_element_by_xpath('//ul[@class="tm-ind-panel"]//li[@class="tm-ind-item tm-ind-reviewCount canClick tm-line3"]').text
        except:item["Cumulative_comments"] =""
        
        # 创建文件夹
        path_new = self.create_file(path,item)

        # 5张主图,下载及命名
        mainFigureCounts = 1
        mainFigure = self.driver.find_elements_by_xpath('//ul[@id="J_UlThumb"]//li')
        for i in mainFigure:
            # 主图链接
            mainPhotoLink = i.find_element_by_xpath('.//img').get_attribute("src").replace('_60x60q90.jpg','')
            # 主图名字
            item["mainFigure{}".format(mainFigureCounts)] = mainPhotoLink.split('/')[-1]
            # 下载图片
            request.urlretrieve(mainPhotoLink,os.path.join(path_new,mainPhotoLink.split('/')[-1]))

            # 调用百度ocr接口做主图ocr
            if ocrSwitch!=None:
                try:
                    item["mainFigure{}_ocr".format(mainFigureCounts)] = self.ocr.picture_Path(os.path.join(path_new,mainPhotoLink.split('/')[-1])) 
                    sleep(0.5)
                except:item["mainFigure{}_ocr".format(mainFigureCounts)] =""
            mainFigureCounts+=1
        sku_content.append(item)
        # SKU列表解析
        sku_list = self.driver.find_elements_by_xpath('//ul[@class="tm-clear J_TSaleProp tb-img     "]//li')
        for i in sku_list:
            item_sku = {}
            # SKU列表点击
            try:
                element=i.find_element_by_xpath('./a')
                self.driver.execute_script("arguments[0].click();",element)
                sleep(0.5)
            except Exception as result:
                print(result)
                return sku_content
            # SKU不为空才加入item
            if i.find_element_by_xpath('./a').text !="":
                # SKU名称
                item_sku["SKU"] = self.eliminateIllegalCharacter(i.find_element_by_xpath('./a').text)
                # SKU价格
                item_sku["Price"] = self.driver.find_elements_by_xpath('//span[@class="tm-price"]')[-1].text
                # SKU图片
                Picture_url = self.driver.find_element_by_xpath('//div[@class="tb-booth"]//img[@id="J_ImgBooth"]').get_attribute("src")
                request.urlretrieve(Picture_url,os.path.join(path_new,item["ID"]+"_"+item_sku["SKU"]+".png"))
            sku_content.append(item_sku)
        #print(sku_content)
        return path_new,sku_content

    def taobao(self,path):
        # 思路和天猫一样，只是页面不一样
        sku_content = []
        # 解析淘宝商品详情页信息
        item = {}
        # 店铺名称
        try:
            item["Shop"] = self.driver.find_element_by_xpath('//div[@class="tb-shop-name"]').text.replace(' ','').replace('\n', '')
        except:
            item["Shop"] = self.driver.find_element_by_xpath('//div[@class="shop-name-wrap"]').text.replace(' ','').replace('\n', '')
        # 商品ID
        item["ID"] = self.parse_id(self.driver.current_url)  
        # 商品重要卖点
        item["keySellingPoint"] = ""
        # 商品标题
        item["NameCommodity"] = self.driver.find_element_by_xpath('//div[@id="J_Title"]/h3').text.strip()
        # 商品类目
        item["category"] = self.driver.find_element_by_xpath('//span[@class="yk-font-red-color yk-other-cat"]').text
        # 产品参数
        item["parameter"] = self.driver.find_element_by_xpath('//ul[@class="attributes-list"]').text.strip().replace("\n","；")
        # 月销量
        item["Monthly_sales"] = self.driver.find_element_by_xpath('//strong[@id="J_SellCounter"]').text
        # 评论数
        item["Cumulative_comments"] = self.driver.find_element_by_xpath('//strong[@id="J_RateCounter"]').text
        
        # 创建文件夹
        path_new = self.create_file(path,item)

        # 5张主图,下载及命名
        mainFigureCounts = 1
        mainFigure = self.driver.find_elements_by_xpath('//ul[@id="J_UlThumb"]//div[@class="tb-pic tb-s50"]')
        for i in mainFigure:
            mainPhotoLink = i.find_element_by_xpath('.//img').get_attribute("src").replace('_50x50.jpg_.webp','')
            item["mainFigure{}".format(mainFigureCounts)] = mainPhotoLink.split('/')[-1]
            request.urlretrieve(mainPhotoLink,os.path.join(path_new,mainPhotoLink.split('/')[-1]))
            if ocrSwitch!=None:
                # 调用百度ocr接口做主图ocr
                try:
                    item["mainFigure{}_ocr".format(mainFigureCounts)] = self.ocr.picture_Path(os.path.join(path_new,mainPhotoLink.split('/')[-1])) 
                    sleep(0.5)
                except:item["mainFigure{}_ocr".format(mainFigureCounts)] =""

            mainFigureCounts+=1
        sku_content.append(item)
        # SKU列表解析
        sku_list = self.driver.find_elements_by_xpath('//ul[@class="J_TSaleProp tb-img tb-clearfix"]//li')
        for i in sku_list:
            item_sku = {}
            # SKU列表点击
            try:
                element=i.find_element_by_xpath('./a')
                self.driver.execute_script("arguments[0].click();",element)
                sleep(0.5)
            except Exception as result:
                print(result)
                return sku_content
            # SKU名称
            item_sku["SKU"] = self.eliminateIllegalCharacter(i.find_element_by_xpath('.//span').get_attribute('innerText'))
            # SKU价格
            item_sku["Price"] = self.driver.find_element_by_xpath('//strong[@id="J_StrPrice"]/em[@class="tb-rmb-num"]').text
            #print(item_sku)
            # SKU图片
            Picture_url = self.driver.find_element_by_xpath('//img[@id="J_ImgBooth"]').get_attribute("src")
            request.urlretrieve(Picture_url,os.path.join(path_new,item["ID"]+"_"+item_sku["SKU"]+".png"))
            sku_content.append(item_sku)
        return path_new,sku_content

    def depositedInExcel(self,path,ocrSwitch,sku_content=None):
        wb = workbook.Workbook()
        sheet1 = wb.active
        cellCount = 2
        sheet1["A1"] = "商品标题"
        sheet1["A2"] = sku_content[0]["NameCommodity"]
        sheet1["B1"] = "重要卖点"
        sheet1["B2"] = sku_content[0]["keySellingPoint"]
        sheet1["C1"] = "商品类目"
        sheet1["C2"] = sku_content[0]["category"]
        sheet1["D1"] = "店铺名称"
        sheet1["D2"] = sku_content[0]["Shop"]
        sheet1["G1"] = "商品ID"
        sheet1["G2"] = sku_content[0]["ID"]
        sheet1["H1"] = "商品参数"
        sheet1["H2"] = sku_content[0]["parameter"]
        sheet1["I1"] = "月销量"
        sheet1["I2"] = sku_content[0]["Monthly_sales"]
        sheet1["J1"] = "累计评价"
        sheet1["J2"] = sku_content[0]["Cumulative_comments"]
        
        # 这个图有的没有
        try:
            sheet1["K1"] = "主图1"
            sheet1["K2"] = sku_content[0]["mainFigure1"]
            sheet1["L1"] = "主图2"
            sheet1["L2"] = sku_content[0]["mainFigure2"]
            sheet1["M1"] = "主图3"
            sheet1["M2"] = sku_content[0]["mainFigure3"]
            sheet1["N1"] = "主图4"
            sheet1["N2"] = sku_content[0]["mainFigure4"]
            sheet1["O1"] = "主图5"
            sheet1["O2"] = sku_content[0]["mainFigure5"]
            if ocrSwitch!=None:
                sheet1["P1"] = "主图1_OCR"
                sheet1["P2"] = sku_content[0]["mainFigure1_ocr"]
                sheet1["Q1"] = "主图2_OCR"
                sheet1["Q2"] = sku_content[0]["mainFigure2_ocr"]
                sheet1["R1"] = "主图3_OCR"
                sheet1["R2"] = sku_content[0]["mainFigure3_ocr"]
                sheet1["S1"] = "主图4_OCR"
                sheet1["S2"] = sku_content[0]["mainFigure4_ocr"]
                sheet1["T1"] = "主图5_OCR"
                sheet1["T2"] = sku_content[0]["mainFigure5_ocr"]
        except:pass
        # SKU需要单独处理
        sheet1["E1"] = "SKU"
        sheet1["F1"] = "SKU价格"
        for i in sku_content[1:]:
            # 解决SKU列表不等于显示SKU列表的问题
            try:
                sheet1["E{}".format(cellCount)] = sku_content[0]["ID"]+"_"+i["SKU"]
                sheet1["F{}".format(cellCount)] = i["Price"]
                cellCount+=1
            except:pass

        path = os.path.join(path,"SKU_{0}_ID{1}_{2}.xlsx".format(time.strftime('%Y-%m-%d'),sku_content[0]["ID"],sku_content[0]["Shop"]))
        # 完成一个表
        print(sku_content[0]["Shop"],sku_content[0]["ID"])
        wb.save(filename=path)

if __name__=="__main__":

    tianmao = TianMao()

    '''
    启动流程：
    --->win + R
    --->输入cmd，回车
    --->复制cd C:\Program Files\Google\Chrome\Application,回车
    --->复制chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenium\AutomationProfile"，回车
    --->关闭黑色命令窗口
    --->在刚才调出的浏览器中登录自己的淘宝天猫账号登录，以后每次调出这个浏览器时检查登录状态即可。
    --->设置用于存储文件夹的路径
    '''

    path = r"D:\Testfiles\cheyongxiche"
    # 这里只要设置ocrSwitch=1,就会启动主图OCR，默认为关闭
    # 在此输入商品ID列表--->Ctrl+s 保存--->shift+alt+F5 启动

    ocrSwitch=None
    beginToList = '''
577333802249
622026790813
618642383637
552181956848
623864406455
627045218144
627330047520
610210528797
597801781421
618732092392
614762967740
556081946184
548901346397
601504376020
626478042299
612735461262
593549642450
614471615733
600282065736
582338152789
620552736199
588774822331
605241786079
618515560261
617596021451
541605779862
527609037982
585602666962
612547300374
592823570185
565882812847
594360709879
620310246000
546102494715
622338914986
586950793089
553824961267
621361502373
590822149739
584884545875
583454649095
591224342685
622415435649
590290233422
619112619588
620962347613
544952583376
591174684987
594746698620
611434717673
619638565483
583564870997
627513921907
627409863424
626077316115
625415498197
626754897136
622434253590
623612756825
538970605154
592685386135
609511385659
622153116640
626717820525
626531964261
589476909573
528038030747
629361434253
630900575707
598776598693
628220126855
621898590660
521078341141
        '''

    tianmao.requestUrlList(beginToList)
    print("-" * 50,"{0}".format(time.strftime('%Y-%m-%d %H:%M')),"-" * 50)