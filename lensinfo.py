import urllib.request
import urllib.error
from bs4 import BeautifulSoup
import os
import re
import xlwt
import sqlite3
import ssl

# 定义正则语句
findlensinfo = re.compile(r'<div class="item-title">.*>(.*?)</a>')
findlensprice = re.compile(r'<div class="price price-now">.*>(.*?)</a>')
# 从lensinfo中拆解详细的焦距、光圈等信息
findcom = re.compile(r'.*?(?=[0-9])')
findmm = re.compile(r'[0-9].*mm')
findff = re.compile(r'mm(.*[0-9])')
findother = re.compile(r'.*[0-9](.*)')


def saveData(datalist, savepath):
    book = xlwt.Workbook(encoding="utf-8")
    sheet = book.add_sheet("sheet1")
    col = ("完整名称","制造商","焦距","光圈","其他型号信息","报价")
    count=len(datalist)
    for i in range(0, 6):
        sheet.write(0, i, col[i])
    for i in range(0, count):
        print("第%d条"%(i+1))
        data=datalist[i]
        for j in range(0,6):
            sheet.write(i+1,j,data[j])
    book.save(savepath)


# 根据基础URL，调用askURL函数，爬取所有相关的网页文件,并以列表返回解析后的数据
def getData(baseurl):
    datalist = []
    for url in baseurl:
        htmlData = askURL(url)  # 调用获取网页数据的函数

        # 解析网页内容
        soup = BeautifulSoup(htmlData, "html.parser")
        for item in soup.find_all('li', class_="item"):
            data = []  # 保存一个镜头的所有信息
            item = str(item)

            lensinfo = re.findall(findlensinfo, item)
            if len(lensinfo) == 0:
                continue

            data.append(lensinfo[0])
            try:
                data.append(re.findall(findcom,lensinfo[0])[0])
            except IndexError:
                data.append("无法匹配到品牌")
                print(lensinfo,"无法匹配到品牌")

            try:
                data.append(re.findall(findmm,lensinfo[0])[0])
            except IndexError:
                data.append("无法匹配到焦距")
                print(lensinfo,"无法匹配到焦距")
            try:
                data.append(re.findall(findff,lensinfo[0])[0])
            except IndexError:
                data.append("无法匹配到光圈")
                print(lensinfo,"无法匹配到光圈")

            try:
                data.append(re.findall(findother,lensinfo[0])[0])
            except IndexError:
                data.append("无法匹配到其他型号信息")
                print(lensinfo,"无法匹配到其他型号信息")

            lensprice = re.findall(findlensprice, item)
            if len(lensprice) != 0:
                data.append(lensprice[0].replace("¥", ""))
            else:
                data.append("")

            datalist.append(data)

    return datalist


# 请求一个URL，将页面内容以字符串的形式返回
def askURL(url):
    head = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.1.1 Safari/605.1.15"
    }
    req = urllib.request.Request(url, headers=head)
    html = ""
    try:
        response = urllib.request.urlopen(req)
        html = response.read().decode("GBK")
        print("获取成功")
    except urllib.error.URLError as e:
        if hasattr(e, "code"):
            print(e.code)
        if hasattr(e, "reason"):
            print(e.reason)
    return html


# def saveHtml(file_name,file_content):
#     #    注意windows文件命名的禁用符，比如 /
#     with open (file_name.replace('/','_'),"w") as f:
#         #   写文件用bytes而不是str，所以要转码
#         f.write( file_content )

if __name__ == '__main__':
    ssl._create_default_https_context = ssl._create_unverified_context

    # 索尼E卡口
    baseurl=["https://product.pconline.com.cn/lens/c11228_c15513/list.shtml","https://product.pconline.com.cn/lens/c11228_c15513/list_25s1.shtml","https://product.pconline.com.cn/lens/c11228_c15513/list_50s1.shtml","https://product.pconline.com.cn/lens/c11228_c15513/list_75s1.shtml"]
    savepath = "./sonylens.xls"

    # M43画幅
    # baseurl=["https://product.pconline.com.cn/lens/c11229/list.shtml","https://product.pconline.com.cn/lens/c11229/list_25s1.shtml"]
    # savepath = "./M43lens.xls"

    datalist = getData(baseurl)
    saveData(datalist, savepath)