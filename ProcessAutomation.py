import pandas as pd
from time import sleep
import time,pyautogui,pyperclip
from openpyxl import load_workbook
from openpyxl import workbook
from openpyxl.utils import FORMULAE
from openpyxl.drawing.image import Image
from PIL import Image
import os,sys,pytesseract,datacompy
from pandas.testing import assert_frame_equal
from retrying import retry
import cv2,requests


def retry_if_io_error(exception):
    # 根据指定异常重试,@retry(retry_on_exception=retry_if_io_error)
    return isinstance(exception, IOError)

def retry_if_TypeError_error(exception):
    # 根据指定异常重试,@retry(retry_on_exception=retry_if_TypeError_error)
    return isinstance(exception, TypeError)

def imageRecognition():
    # 图像识别
    path = r"C:\Users\BJB00936\Desktop\钢化膜.jpg"
    image = Image.open(path)
    text = pytesseract.image_to_string(image,lang="chi_sim+eng",config="--psm 6 --oem 3 -c tessedit_char_whitelist=0123456789")
    print(text)

def cancelCompetingGoods():
    # 取消竞品配置
    while True:
        pyautogui.click(1591,645)
        pyautogui.click(1107,299)
        sleep(1)

def competitiveProductConfiguration():
    # 竞品配置，添加转化缓存
    df = pd.read_excel(r"G:\testfiles\市场分析_声卡 直播专用.xlsx",sheet_name="市场概况源数据")
    df.sort_values(by='30天销量',ascending=False,na_position='last')
    for i in df["商品id"]:
        # 点击 "+"
        pyautogui.click(577,523)
        # 点击输入框
        pyautogui.click(574,578)
        # 输入产品id
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(i)
        pyautogui.hotkey('ctrl', 'v')
        sleep(1)
        # 点击对应竞品
        pyautogui.click(581,613)
        sleep(2.5)
        # 截图
        pyautogui.screenshot('{}.jpg'.format(os.path.join(r"G:\testfiles",str(i))),(432,582,1201,353))
        pyautogui.click(700,505)
        sleep(0.5)
        print(i)

def addCompetingGoods():
    # 添加竞品
    df = pd.read_excel(r"G:\testfiles\test.xlsx",sheet_name="Sheet1")
    for i in df["商品id"]:
        # 点击 "+"
        pyautogui.click(577,435)
        # 点击输入框
        pyautogui.click(574,494)
        # 输入产品id
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(i)
        pyautogui.hotkey('ctrl', 'v')
        sleep(1)
        # 点击对应竞品
        pyautogui.click(581,520)
        # 添加监控
        pyautogui.click(1569,439)
        sleep(1.5)
        # 确认监控
        pyautogui.click(1445,455)
        sleep(0.5)

def downloadMarket():
    # 下载大盘
    df = pd.read_excel(r"D:\Testfiles\text.xlsx",sheet_name="Sheet1")
    for i in df["类目"]:
        # 点击 "类目"
        pyautogui.click(594,166)
        # 点击输入框
        pyautogui.click(561,198)
        # 输入产品id
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(i)
        pyautogui.hotkey('ctrl', 'v')
        sleep(1)
        # 点击对应类目
        pyautogui.click(569,231)
        # 点击一键转化
        pyautogui.click(1261,343)
        sleep(1)
        # 添加加载缓存
        pyautogui.click(272,255)
        pyautogui.click(412,255)
        pyautogui.click(569,255)
        pyautogui.click(733,255)
        pyautogui.click(877,255)
        pyautogui.click(1027,255)
        pyautogui.click(1189,255)
        pyautogui.click(1341,255)
        pyautogui.click(1470,255)
        pyautogui.click(1644,255)
        sleep(1)
        # 导出csv
        pyautogui.click(1612,550)
        sleep(1)
        # 点击保存
        pyautogui.click(1187,781)
        # 关闭转化
        pyautogui.click(1743,150)
        sleep(0.5)

def downloadBrand():
    # 下载品牌
    df = pd.read_excel(r"D:\Testfiles\text.xlsx",sheet_name="Sheet1")
    for i in df["类目"]:
        # 点击月
        pyautogui.click(1495,231)
        sleep(3)
        # 点击 "类目名"
        pyautogui.click(586,228)
        sleep(1)
        # 点击输入框
        pyautogui.click(571,260)
        # 输入类目名
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(i)
        pyautogui.hotkey('ctrl', 'v')
        sleep(5)
        # 点击对应类目
        pyautogui.click(572,291)
        sleep(1)
        # 点击品牌
        pyautogui.click(643,279)
        for i2 in range(0,24):
            # 点击高交易
            pyautogui.click(613,334)
            #sleep(2)
            ## 点击高流量
            #pyautogui.click(666,333)
            sleep(2)
            ## 点击合并转化
            #pyautogui.click(1054,277)
            # 点击一键转化
            pyautogui.click(963,277)
            sleep(4)
            # 导出csv
            pyautogui.click(1311,445)
            sleep(2)
            # 点击保存
            pyautogui.click(600,445)
            # 关闭转化
            pyautogui.click(157,595)
            # 上个月
            pyautogui.click(1525,230)
            sleep(3)

@retry(wait_fixed=2000)
def aKeyTransformation():
    # 点击一键转化
    x,y = imageCenter = pyautogui.center(pyautogui.locateOnScreen(aKey_Transformation_img, confidence=0.8))
    pyautogui.click(x,y)

@retry(wait_fixed=2000)
def exportCSV():
    # 导出csv
    x,y = imageCenter = pyautogui.center(pyautogui.locateOnScreen(export_CSV_img, confidence=0.8))
    pyautogui.click(x,y)

@retry(wait_fixed=2000)
def confirmThatSavePath():
    # 点击保存
    x,y = imageCenter = pyautogui.center(pyautogui.locateOnScreen(save_CSV_img, confidence=0.8))
    pyautogui.click(x,y)

@retry(wait_fixed=2000)
def closeTransformation():
    # 关闭转化
    pyautogui.click(157,595)

@retry(wait_fixed=2000)
def cycleBackButton():
    # 时间周期回退按钮
    x,y = pyautogui.center(pyautogui.locateOnScreen(cycle_Back_img, confidence=0.8))
    pyautogui.click(x,y)

@retry(wait_fixed=2000)
def resetTime():
    # 时间重置
    x,y = pyautogui.center(pyautogui.locateOnScreen(month_img, confidence=0.8))
    pyautogui.click(x,y)

@retry(wait_fixed=2000)
def clickOnCategory():
    # 点击类目
    x,y = pyautogui.center(pyautogui.locateOnScreen(enterCategory_img, confidence=0.8))
    pyautogui.click(x,y)

#@retry(wait_fixed=2000)
#def clickOnCategory():
#    # 点击类目
#    x,y = pyautogui.center(pyautogui.locateOnScreen(enterCategory_img, confidence=0.8))
#    pyautogui.click(x,y)

@retry(wait_fixed=2000)
def clickOnCategoryInputBox():
    # 点击类目输入框
    x,y = pyautogui.center(pyautogui.locateOnScreen(enterCategory_img, confidence=0.8))
    pyautogui.click(x,y+30)

@retry(wait_fixed=2000)
def chooseCategory():
    # 选择类目，这个有点问题
    x,y = pyautogui.center(pyautogui.locateOnScreen(enterCategory_img, confidence=0.8))
    pyautogui.click(x,y+60)

@retry(wait_fixed=2000)
def clickOnBrand():
    # 点击品牌
    x,y = pyautogui.center(pyautogui.locateOnScreen(enterCategory_img, confidence=0.8))
    pyautogui.click(x,y+50)


def highBrand():
    # 商品排行高交易csv下载
    #df = pd.read_excel(r"D:\Testfiles\text.xlsx",sheet_name="Sheet1")
    #for i in df["类目"]:
    #    # 点击 "类目"
    #    pyautogui.click(594,166)
    #    # 点击输入框
    #    pyautogui.click(561,198)
    #    # 输入产品id
    #    pyautogui.hotkey('ctrl', 'a')
    #    pyperclip.copy(i)
    #    pyautogui.hotkey('ctrl', 'v')
    #    sleep(1)
    #    # 点击对应类目
    #    pyautogui.click(569,231)
    #    # 点击一键转化
    #    pyautogui.click(1261,343)

    for i2 in range(0,7):
        sleep(2)
        # 一键转化
        aKeyTransformation()
        sleep(6)
        # 导出csv
        exportCSV()
        sleep(2)
        # 关闭转化
        closeTransformation()
        sleep(2)
        # 时间回退
        cycleBackButton()
        sleep(3)
    ## 点击月
    #resetTime()
    #sleep(2)
    ## 点击品牌
    #clickOnBrand()
    #sleep(3)
    #for i2 in range(0,11):
    #    sleep(2)
    #    # 一键转化
    #    aKeyTransformation()
    #    sleep(6)
    #    # 导出csv
    #    exportCSV()
    #    sleep(2)
    #    # 关闭转化
    #    closeTransformation()
    #    sleep(2)
    #    # 时间回退
    #    cycleBackButton()
    #    sleep(3)

def downloadIndustryCustomers():
    # 下载行业客群
    df = pd.read_excel(r"D:\Testfiles\text.xlsx",sheet_name="Sheet1")
    for i in df["类目"]:
        # 点击 "类目名"
        pyautogui.click(588,231)
        sleep(1)
        # 点击输入框
        pyautogui.click(584,259)
        # 输入类目名
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(i)
        pyautogui.hotkey('ctrl', 'v')
        sleep(5)
        # 点击对应类目
        pyautogui.click(585,294)
        sleep(3)
        # 点击客群占比、交易指数等
        pyautogui.click(495,793)
        sleep(1)
        pyautogui.click(570,791)
        sleep(1)
        pyautogui.click(647,791)
        sleep(1)
        pyautogui.click(723,792)
        sleep(1)
        # 一键转化
        pyautogui.click(1431,742)
        sleep(4)
        # 导出csv
        pyautogui.click(1427,542)
        sleep(2)
        # 点击保存
        pyautogui.click(602,445)
        # 关闭转化
        pyautogui.click(153,596)
        sleep(2)
        pyautogui.hotkey('ctrl', 'shift','y')
        sleep(5)

def downloadIndustryCustomers():
    # 下载行业客群透视
    df = pd.read_excel(r"D:\Testfiles\text.xlsx",sheet_name="Sheet1")
    for i in df["类目"]:
        # 点击 "类目名"
        pyautogui.click(592,231)
        sleep(1)
        # 点击输入框
        pyautogui.click(558,260)
        # 输入类目名
        pyautogui.hotkey('ctrl', 'a')
        pyperclip.copy(i)
        pyautogui.hotkey('ctrl', 'v')
        sleep(5)
        # 点击对应类目
        pyautogui.click(566,291)
        sleep(3)
        # 点击客群占比、交易指数等
        pyautogui.click(555,371)
        sleep(1)
        pyautogui.click(646,371)
        sleep(1)
        # 一键转化
        pyautogui.click(1424,323)
        sleep(4)
        # 导出csv
        pyautogui.click(1410,223)
        sleep(2)
        # 点击保存
        pyautogui.click(602,445)
        # 关闭转化
        pyautogui.click(153,596)
        sleep(2)
        pyautogui.hotkey('ctrl', 'shift','y')
        sleep(5)

def findPositionAccordingToPictures():
    # 测试图片定位，加上OpenCV后效果好
    try:
        x,y = imageCenter = pyautogui.center(pyautogui.locateOnScreen(headsetHeadset_img, confidence=0.8))
        print(x,y)
    except TypeError as error1:print("定位不到",error1)


pyautogui.PAUSE=2  # 基本停止
pyautogui.FAILSAFE = True  # 错误停止
#生意参谋系
aKey_Transformation_img = r"D:\Testfiles\ImagePosition\yijianzhuanhua.jpg"
export_CSV_img = r"D:\Testfiles\ImagePosition\daochucsv.jpg"
save_CSV_img = r"D:\Testfiles\ImagePosition\baocun.jpg"
cycle_Back_img = r"D:\Testfiles\ImagePosition\zhouqidaotui.jpg"
month_img = r"D:\Testfiles\ImagePosition\yue.jpg"
brand_img = r"D:\Testfiles\ImagePosition\pinpai.jpg"
goods_img = r"D:\Testfiles\ImagePosition\shangpin.jpg"
enterCategory_img = r"D:\Testfiles\ImagePosition\zhuanyeban.jpg"

#网页阿明工具系
review_Analysis_img = r"D:\Testfiles\ImagePosition\reviewAnalysis.jpg"
aming_daochuCSV_img = r"D:\Testfiles\ImagePosition\daochuCSV.jpg"
    

if __name__=="__main__":

    #findPositionAccordingToPictures()  # 根据图片识别屏幕位置

    #imageRecognition()  # OCR

    #addCompetingGoods()  # 添加竞品配置

    #downloadMarket()  # 下载大盘

    #downloadBrand()  # 下载品牌

    highBrand()  # 商品&品牌排行高交易csv下载

    #downloadIndustryCustomers()  # 行业客群下载并截图

    print("-" * 50,"{0}".format(time.strftime('%Y-%m-%d %H:%M')),"-" * 50)