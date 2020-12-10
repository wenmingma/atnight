import time
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

class TianMao:

    def __init__(self):
        
        #cd C:\Program Files (x86)\Google\Chrome\Application
        #chrome.exe --remote-debugging-port=9222 --user-data-dir="E:\selenium\AutomationProfile"
        options = Options()
        options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        chrome_driver = r"C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe" 
        self.driver = webdriver.Chrome(chrome_driver,options=options)
        script = '''
        Object.defineProperty(navigator, 'webdriver', {
        get: () => undefined
        })
        '''
        self.driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": script})

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
            x1 = x["data-start"].replace(".00","").replace(".0","")
            if x["data-end"]!="":
                x2 = x["data-end"].replace(".00","")
            else:TaobaoHomepageUsersLove
                x2 = ""
            print("{}：{}-{}".format(x["aria-label"],x1,x2))

    def J_TAllCatsTree_cats_tree(self):
        # 抓取商家店铺首页分类
        title = self.driver.title.strip("首页--Tmall.com")
        cats_trees = self.driver.find_elements_by_xpath('//li[@class="cat fst-cat"]')
        for cats_tree in cats_trees:
            x_list = []
            #i1 = cats_tree.find_element_by_xpath('.//*').get_attribute('textContent').replace(' ','').replace('\n', '')
            i1 = cats_tree.find_element_by_xpath('./*').get_attribute('textContent').replace(' ','').replace('\n', '')
            i2 = cats_tree.find_elements_by_xpath('.//li[@class="cat snd-cat"]')
            for x in i2:
                x_content = x.find_element_by_xpath('.').get_attribute('textContent').replace(' ','').replace('\n', '')
                x_list.append(x_content)
            for x2 in x_list:
                print(title,i1,x2)


        # 获取隐藏元素，鼠标悬浮才弹出文本的方法
        #print(cats_tree.get_attribute('innerHTML'))
        #print(self.driver.execute_script("return arguments[0].innerHTML",cats_tree))
        #print(cats_tree.get_attribute('textContent'))
        #print(self.driver.execute_script("return arguments[0].textContent",cats_tree))

    def run(self):
        print(self.driver.title)
        self.J_SortbarPriceRanks()
        #self.J_TAllCatsTree_cats_tree()


if __name__=="__main__":
    tianmao = TianMao()
    tianmao.run()
    print("-" * 50,"{0}".format(time.strftime('%Y-%m-%d %H:%M')),"-" * 50)