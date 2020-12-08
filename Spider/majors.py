import re
import requests
from lxml import etree
from openpyxl import Workbook

id = 1

def newfile():
    wb = Workbook()
    ws = wb.active
    ws.title = "专业信息"
    titlelist = ["序号", "院校名称", "硕士/博士专业", "专业名称", "方向名称", "方向代码"]
    ws.append(titlelist)
    return wb

def get_xpath(url):
    heads = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"}
    try:
        response = requests.get(url=url, headers=heads)
        return etree.HTML(response.text)
    except Exception:
        print(url, '该页面没有相应！')
        return ''

def get_basic_facts(item,wb):
    global id
    ws = wb.active
    name = item.xpath('./td[1]/a/text()')
    name = re.sub(r'\s+', "", name[0])
    surl = item.xpath('./td[1]/a/@href')
    surlc = 'https://yz.chsi.com.cn' + surl[0]
    html = get_xpath(surlc)
    schoolLink = html.xpath('//ul[@class="yxk-link-list clearfix"]')
    curl = schoolLink[0].xpath('./li[3]/a/@href')
    curlc = 'https://yz.chsi.com.cn' + curl[0]
    zy_0 = get_xpath(curlc).xpath('//div[@class="container"]/div[2]/div[@class="ch-tab clearfix"]/div/a/text()')
    zy_1 = get_xpath(curlc).xpath('//div[@class="container"]/div[2]/div[2]/div/ul')
    s = 0
    print(name)
    print('硕士专业')
    for item in zy_1:
        for x in item.xpath('./li'):
            lx = eval(str(x.xpath('.//text()')).replace(r'\r\n', '').replace(' ', ''))
            # print(lx)
            if lx[0]:
                i = lx[0]
            else:
                i = lx[1]
            dataList = [id]
            id = id + 1
            dataList.append(name)
            dataList.append('硕士')
            dataList.append(zy_0[s])
            dataList.append(i[:-8])
            dataList.append(i[-7:-1])
            ws.append(dataList)
            print(dataList)
        s += 1
    zy_2 = get_xpath(curlc).xpath('//div[@class="container"]/div[4]/div[@class="ch-tab clearfix"]/div/a/text()')
    zy_3 = get_xpath(curlc).xpath('//div[@class="container"]/div[4]/div[2]/div/ul')
    ss = 0
    print('博士专业')
    for item in zy_3:
        for x in item.xpath('./li'):
            lx = eval(str(x.xpath('.//text()')).replace(r'\r\n', '').replace(' ', ''))
            # print(lx)
            if lx[0]:
                i = lx[0]
            else:
                i = lx[1]
            dataList = [id]
            id = id + 1
            dataList.append(name)
            dataList.append('博士')
            dataList.append(zy_2[ss])
            dataList.append(i[:-8])
            dataList.append(i[-7:-1])
            print(dataList)
            ws.append(dataList)
        ss += 1
    wb.save("专业信息.xlsx")



def all_data(url,file):
    collegeList = get_xpath(url).xpath('//div[@class="yxk-table"]')
    for item in collegeList:
        dataList = item.xpath('./table/tbody/tr')
    for item in dataList:
        get_basic_facts(item,file)



def main():
    url = "https://yz.chsi.com.cn/sch/search.do?ssdm=51&start="
    file = newfile()
    for i in range(0,2):
        all_data(url+str(i*20),file)


if __name__ =="__main__":
    main()