from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from lxml import etree
from time import sleep
from decimal import Decimal
import random,csv


class TianMao:

    def __init__(self):

        #cd C:\Program Files\Google\Chrome\Application
        #chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\selenium\AutomationProfile"
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
        print(self.driver.title)

    def roll(self):

        js="var q=document.documentElement.scrollTop=10000"
        self.driver.execute_script(js)
        sleep(1)

    def remove_exponent(self,num):
        #去掉数字无效的0
        return num.to_integral() if num == num.to_integral() else num.normalize()

    def J_SortbarPriceRanks(self):
        # 抓取页面价格区间，%几用户喜欢多少多少
        PriceRanks = self.driver.find_elements_by_xpath('//div[@class="items J_SortbarPriceRanks"]//a')
        contents = []
        for i in PriceRanks:
            item = {}
            item["data-start"] = i.find_element_by_xpath('.').get_attribute("data-start")
            item["aria-label"] = i.find_element_by_xpath('.').get_attribute("aria-label")
            item["data-end"] = i.find_element_by_xpath('.').get_attribute("data-end")
            contents.append(item)
        for x in contents:
            x1 = x["data-start"]
            if x["data-end"]!="":
                x2 = x["data-end"]
            else:
                x2 = ""
            print("{}：{}-{}".format(x["aria-label"],x1,x2))



    def run(self):
        print(self.driver.title)
        self.J_SortbarPriceRanks()


if __name__=="__main__":
    tianmao = TianMao()
    tianmao.run()
    print("-"*50,"succeed","-"*50)