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
        # 解析beginToList并发送请求
        for i in beginToList.strip().split():
            # 请求URL
            if str(i).find("html")==-1:
                self.driver.get(self.linksToCompletion(i))
            else:
                self.driver.get(i)
            sleep(4)
            pyautogui.scroll(50)
            try:
                # 抓取京东页面信息
                self.jingdong(path)
            except Exception as error1:
                print("商品可能下架或预售期：{}，或出现{}".format(i,error1))

    def jingdong(self,path):

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
        
        print(item)
        
        
        
       
        
        
        
    #    # 创建文件夹
    #    path_new = self.create_file(path,item)

    #    # 5张主图,下载及命名
    #    mainFigureCounts = 1
    #    mainFigure = self.driver.find_elements_by_xpath('//ul[@id="J_UlThumb"]//li')
    #    for i in mainFigure:
    #        mainPhotoLink = i.find_element_by_xpath('.//img').get_attribute("src").replace('_60x60q90.jpg','')
    #        item["mainFigure{}".format(mainFigureCounts)] = mainPhotoLink.split('/')[-1]
    #        r = requests.get(mainPhotoLink)
    #        with open(os.path.join(path_new,mainPhotoLink.split('/')[-1]),'wb') as f:
    #            f.write(r.content)

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

    #def create_file(self,path,item):
    ## 判断文件夹是否已经存在，若存在则跳过，不存在则创建
    #    inspection_path = os.path.join(path,"{0}_ID{1}_{2}".format(time.strftime('%Y-%m-%d-%H'),item["ID"],item["Shop"]))
    #    if not os.path.exists(inspection_path):
    #        os.makedirs(os.path.join(path,"{0}_ID{1}_{2}".format(time.strftime('%Y-%m-%d-%H'),item["ID"],item["Shop"])))
    #    return inspection_path

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
https://item.jd.com/100003653795.html
https://item.jd.com/2672959.html
https://item.jd.com/100014740254.html
https://item.jd.com/100011105006.html
https://item.jd.com/35445129381.html
https://item.jd.com/100000006775.html
https://item.jd.com/50990950435.html
https://item.jd.com/100008347304.html
https://item.jd.com/1459061.html
https://item.jd.com/100013127648.html
https://item.jd.com/1377604.html
https://item.jd.com/100006152353.html
https://item.jd.com/10021815738590.html
https://item.jd.com/100007056406.html
https://item.jd.com/100004619429.html
https://item.jd.com/70915175232.html
https://item.jd.com/69387601808.html
https://item.jd.com/3768166.html
https://item.jd.com/50450268458.html
https://item.jd.com/5148412.html
https://item.jd.com/10022029436029.html
https://item.jd.com/10258426739.html
https://item.jd.com/6379895.html
https://item.jd.com/10461744649.html
https://item.jd.com/6181922.html
https://item.jd.com/100000827222.html
https://item.jd.com/100006578035.html
https://item.jd.com/1263417213.html
https://item.jd.com/100007382724.html
https://item.jd.com/11141148463.html
https://item.jd.com/100003587559.html
https://item.jd.com/18860484745.html
https://item.jd.com/28142190977.html
https://item.jd.com/54550692716.html
https://item.jd.com/64661456440.html
https://item.jd.com/100013676518.html
https://item.jd.com/58588750842.html
https://item.jd.com/100013087236.html
https://item.jd.com/43758373911.html
https://item.jd.com/45835830963.html
https://item.jd.com/100006102471.html
https://item.jd.com/58423962403.html
https://item.jd.com/100004246598.html
https://item.jd.com/28772217872.html
https://item.jd.com/100003613228.html
https://item.jd.com/41412027628.html
https://item.jd.com/70402225873.html
https://item.jd.com/55726975717.html
https://item.jd.com/3626464.html
https://item.jd.com/52587864254.html
https://item.jd.com/52692330229.html
https://item.jd.com/53627070231.html
https://item.jd.com/3220971.html
https://item.jd.com/7070804.html
https://item.jd.com/70196296418.html
https://item.jd.com/68407477877.html
https://item.jd.com/8713562.html
https://item.jd.com/57993689770.html
https://item.jd.com/29912805359.html
https://item.jd.com/5731372.html
https://item.jd.com/3901916.html
https://item.jd.com/62396274756.html
https://item.jd.com/10448010616.html
https://item.jd.com/7281349.html
https://item.jd.com/5934006.html
https://item.jd.com/1260578.html
https://item.jd.com/65008485128.html
https://item.jd.com/10021629726444.html
https://item.jd.com/32111086646.html
https://item.jd.com/10021713335414.html
https://item.jd.com/100013351546.html
https://item.jd.com/3569675.html
https://item.jd.com/31543363123.html
https://item.jd.com/8563165.html
https://item.jd.com/13852070992.html
https://item.jd.com/100016406328.html
https://item.jd.com/5565875.html
https://item.jd.com/29234500785.html
https://item.jd.com/66819790087.html
https://item.jd.com/1272817.html
https://item.jd.com/41409103836.html
https://item.jd.com/40196878225.html
https://item.jd.com/28318833092.html
https://item.jd.com/37847213329.html
https://item.jd.com/30677876925.html
https://item.jd.com/10157004703.html
https://item.jd.com/55523407976.html
https://item.jd.com/13206225909.html
https://item.jd.com/52124271017.html
https://item.jd.com/5565879.html
https://item.jd.com/6424567.html
https://item.jd.com/11253879050.html
https://item.jd.com/1272816.html
https://item.jd.com/6128490.html
https://item.jd.com/17128630746.html
https://item.jd.com/29344507195.html
https://item.jd.com/33793220609.html
https://item.jd.com/50496243122.html
https://item.jd.com/29551936858.html
https://item.jd.com/10379458061.html
https://item.jd.com/100009428814.html
https://item.jd.com/10122725781.html
https://item.jd.com/48807140959.html
https://item.jd.com/37847189565.html
https://item.jd.com/46545749756.html
https://item.jd.com/71494494040.html
https://item.jd.com/18248899321.html
https://item.jd.com/35024134303.html
https://item.jd.com/48477763644.html
https://item.jd.com/32514504456.html
https://item.jd.com/1075908019.html
https://item.jd.com/10021561889266.html
https://item.jd.com/10540497594.html
https://item.jd.com/55846838815.html
https://item.jd.com/10356115444.html
https://item.jd.com/12489947748.html
https://item.jd.com/12309052250.html
https://item.jd.com/37847163087.html
https://item.jd.com/50474519185.html
https://item.jd.com/8924505.html
https://item.jd.com/10025036618973.html
https://item.jd.com/37847163052.html
https://item.jd.com/37847208866.html
https://item.jd.com/100012825260.html
https://item.jd.com/10375115145.html
https://item.jd.com/43194622732.html
https://item.jd.com/50385785699.html
https://item.jd.com/100002948370.html
https://item.jd.com/100009797438.html
https://item.jd.com/12899688617.html
https://item.jd.com/49855083018.html
https://item.jd.com/31509983438.html
https://item.jd.com/31508507286.html
https://item.jd.com/37847195075.html
https://item.jd.com/10121967545.html
https://item.jd.com/34249690019.html
https://item.jd.com/29550715494.html
https://item.jd.com/42167348314.html
https://item.jd.com/11218117466.html
https://item.jd.com/1465760656.html
https://item.jd.com/70622190922.html
https://item.jd.com/12224478823.html
https://item.jd.com/27061747633.html
https://item.jd.com/13537317520.html
https://item.jd.com/52762414296.html
https://item.jd.com/44094398653.html
https://item.jd.com/43212121369.html
https://item.jd.com/41408099315.html
https://item.jd.com/100005596927.html
https://item.jd.com/1612376703.html
https://item.jd.com/100007283358.html
https://item.jd.com/20599490217.html
https://item.jd.com/10471334118.html
https://item.jd.com/59253811048.html
https://item.jd.com/1780112155.html
https://item.jd.com/41408090405.html
https://item.jd.com/27826693407.html
https://item.jd.com/5894776.html
https://item.jd.com/42354978834.html
https://item.jd.com/34323068363.html
https://item.jd.com/5809195.html
https://item.jd.com/1567158661.html
https://item.jd.com/26520664379.html
https://item.jd.com/1561750265.html
https://item.jd.com/100012820098.html
https://item.jd.com/1272802.html
https://item.jd.com/53583139484.html
https://item.jd.com/100000006801.html
https://item.jd.com/17256621594.html
https://item.jd.com/5894636.html
https://item.jd.com/100008631249.html
https://item.jd.com/10282899427.html
https://item.jd.com/100015766988.html
https://item.jd.com/1511783665.html
https://item.jd.com/32214082564.html
https://item.jd.com/13187298385.html
https://item.jd.com/1567159953.html
https://item.jd.com/1583879457.html
https://item.jd.com/1164101595.html
https://item.jd.com/11729530559.html
https://item.jd.com/68627972384.html
https://item.jd.com/25323253213.html
https://item.jd.com/10021462100315.html
https://item.jd.com/33205377078.html
https://item.jd.com/69586433905.html
https://item.jd.com/28022747951.html
https://item.jd.com/65624952394.html
https://item.jd.com/1267947.html
https://item.jd.com/11906609438.html
https://item.jd.com/11944616327.html
https://item.jd.com/4024470.html
https://item.jd.com/17242804845.html
        '''
    jd.requestUrlList(beginToList)
    print("-" * 50,"{0}".format(time.strftime('%Y-%m-%d %H:%M')),"-" * 50)