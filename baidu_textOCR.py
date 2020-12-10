from aip import AipOcr
import os,sys,time,jsonpath,json
from pprint import pprint
from openpyxl import load_workbook
from openpyxl import workbook
from openpyxl.utils import FORMULAE
from openpyxl.drawing.image import Image



class BaiduOCR:

    def __init__(self):
        APP_ID = '23091405'
        API_KEY = '8iGi4BkBWNt0yLzBkyADOufI'
        SECRET_KEY = 'a30Rkk6sYTYq0GT4aOGoooQlkSn82uhz'
        self.client = AipOcr(APP_ID, API_KEY, SECRET_KEY)

    """ 读取图片 """
    def get_file_content(self,filePath):
        with open(filePath, 'rb') as fp:
            return fp.read()

    def picture_Path(self,picture_path):
        image = self.get_file_content(picture_path)

        """ 调用通用文字识别（高精度版） """
        self.client.basicAccurate(image)
        """ 调用通用文字识别（标准版） """
        #self.client.basicGeneral(image)

        """ 如果有可选参数 """
        options = {}
        options["detect_direction"] = "true"
        options["probability"] = "true"

        """ 带参数调用通用文字识别（高精度版） """
        generationAnalytical = self.client.basicAccurate(image, options)
        #print(generationAnalytical)
        identifyText = []
        for value in generationAnalytical['words_result']:
            #print(value["words"])
            identifyText.append(value["words"])
        
        identifyTextStr = "_".join(identifyText)
        return identifyTextStr

def toGetImagePath(path):

    wb = workbook.Workbook()
    sheet1 = wb.active
    
    # 第一个表头：文件ID+主图名称
    sheet1["A1"] = "图片信息"
    sheet1["B1"] = "图片"
    sheet1["C1"] = "OCR"

    # 拿到图片数量
    filecount = 2
    for root,dirs,files in os.walk(path):
        filecount+=len(files)

        # 拿到图片路径
        for root,dirs,files in os.walk(path):
            # 拿到图片路径
            for name in files:
                # 拿到图片路径
                imagePath = os.path.join(root,name)
                #print(imagePath)

                # 文件id标识+文件名=商品信息
                commodityInformation = root.split("\\")[-1].split("_")[0] + os.path.splitext(name)[0]
                #print(commodityInformation)

                # 第一列插入商品信息
                sheet1["A{}".format(filecount)] = commodityInformation
                # 第二列插入图片，首先调整行高和列宽
                sheet1.column_dimensions['B'].width = 80
                sheet1.row_dimensions[filecount].height = 80
                img = Image(imagePath)
                # 生成新的图片的宽和高
                newsize = (160,160)
                img.width,img.height = newsize
                # 将图片添加到excel中
                sheet1.add_image(img,'B{}'.format(filecount))
                # 第三列为OCR结果
                sheet1["C{}".format(filecount)] = ocr.picture_Path(imagePath)
                filecount+=1

    wb.save(filename=os.path.join(r"D:\atnight\蓝禾科技","result.xlsx"))


if __name__=="__main__":

    ocr = BaiduOCR()

    path = r"D:\atnight\蓝禾科技\毛球修剪器TOP10-PC主图\毛球修剪器TOP10-PC主图"

    toGetImagePath(path)


    print("-" * 50,"{0}".format(time.strftime('%Y-%m-%d %H:s%M')),"-" * 50)

