import re
import time

import requests
from lxml import etree
from openpyxl import Workbook

id = 1

def newfile():
    wb = Workbook()
    ws = wb.active
    ws.title = "资讯信息"
    titlelist = ["序号","院校名称","咨询类别","标题","问题","回答","时间","院系"]
    ws.append(titlelist)
    return wb

def str_transform_list(str_list):
    list_str = list(str_list)
    return list_str

def list_transfrom_str(list_str):
    str_list = ''.join(list_str)
    return str_list

def data_transfrom(time):
    data = str_transform_list(time)
    data.insert(3, " ")
    data.insert(7, " ")
    data.insert(10, " ")
    data.insert(19, " ")
    data.insert(23, " ")
    return list_transfrom_str(data)

def get_tlq(name,flag,html_test,wb):
    global id
    ws = wb.active
    for item in range(flag * 2 + 1, len(html_test.xpath('//tr[@class="question_cnt_tr"]')) * 2, 2):
        dataList = [id]
        title = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[2]/a/text()')).replace(r'\r\n', '').replace(' ', ''))
        department = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[3]/div/text()')).replace(r'\r\n', '').replace(' ', ''))
        cst_time = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[@class="question_t ch-table-center"]/text()')).replace(r'\r\n', '').replace(' ', ''))
        time = get_time(cst_time[0])
        question = eval(
            str(html_test.xpath('//tr[' + str(item + 1) + ']/td[2]/div/div[1]/text()')).replace(r'\r\n', '').replace(
                ' ', ''))
        anwser = eval(
            str(html_test.xpath('//tr[' + str(item + 1) + ']/td[2]/div/div[2]//text()')).replace(r'\r\n', '').replace(
                ' ', ''))
        del anwser[1]
        anwser = ''.join(anwser)
        dataList.append(name)
        dataList.append("讨论区")
        dataList.append(title[0])
        dataList.append(question[0])
        dataList.append(anwser)
        dataList.append(time)
        dataList.append(department[0])
        id += 1
        ws.append(dataList)
        print(dataList)
def get_ggq(name,html_test,wb):
    global id
    ws = wb.active
    for item in range(1, len(html_test.xpath('//tr[@class="question_cnt_tr"]')) * 2, 2):
        dataList = [id]
        title = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[2]/a/text()')).replace(r'\r\n', '').replace(' ', ''))
        department = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[3]/div/text()')).replace(r'\r\n', '').replace(' ', ''))
        cst_time = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[@class="question_t ch-table-center"]/text()')).replace(r'\r\n', '').replace(' ', ''))
        time = get_time(cst_time[0])
        question = title[0]
        anwser = eval(
            str(html_test.xpath('//tr[' + str(item + 1) + ']/td[2]/div/div//text()')).replace(r'\r\n', '').replace(
                ' ', ''))
        anwser = ''.join(anwser)
        dataList.append(name)
        dataList.append("公告区")
        dataList.append(title[0])
        dataList.append(question)
        dataList.append(anwser)
        dataList.append(time)
        dataList.append(department[0])
        id += 1
        ws.append(dataList)
        print(dataList)
def get_jhq(name,html_test,wb):
    global id
    ws = wb.active
    for item in range(1, len(html_test.xpath('//tr[@class="question_cnt_tr"]')) * 2, 2):
        dataList = [id]
        title = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[2]/a/text()')).replace(r'\r\n', '').replace(' ', ''))
        department = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[3]/div/text()')).replace(r'\r\n', '').replace(' ', ''))
        cst_time = eval(
            str(html_test.xpath('//tr[' + str(item) + ']/td[5]/text()')).replace(r'\r\n', '').replace(' ', ''))
        time = get_time(cst_time[0])
        question = eval(
            str(html_test.xpath('//tr[' + str(item + 1) + ']/td[2]/div/div[1]/text()')).replace(r'\r\n', '').replace(
                ' ', ''))
        anwser = eval(
            str(html_test.xpath('//tr[' + str(item + 1) + ']/td[2]/div/div[2]//text()')).replace(r'\r\n', '').replace(
                ' ', ''))
        del anwser[1]
        anwser = ''.join(anwser)
        dataList.append(name)
        dataList.append("精华区")
        dataList.append(title[0])
        dataList.append(question[0])
        dataList.append(anwser)
        dataList.append(time)
        dataList.append(department[0])
        id += 1
        ws.append(dataList)
        print(dataList)

def trans_format(time_string, from_format, to_format='%Y.%m.%d %H:%M:%S'):
    """
    @note 时间格式转化
    :param time_string:
    :param from_format:
    :param to_format:
    :return:
    """
    time_struct = time.strptime(time_string,from_format)
    times = time.strftime(to_format, time_struct)
    return times

def get_time(data):
    data = data_transfrom(data)
    format_time = trans_format(data, '%a %b %d %H:%M:%S CST %Y', '%Y-%m-%d %H:%M:%S')
    return format_time

def get_xpath(url):
    heads = {
        "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.198 Safari/537.36"}
    try:
        response = requests.get(url=url,headers=heads)
        return etree.HTML(response.text)
    except Exception:
        print(url, '该页面没有相应！')
        return ''

def get_basic_facts(item,wb):
    global id
    url = "https://yz.chsi.com.cn"
    name = item.xpath('./td[1]/a/text()')
    name = re.sub(r'\s+', "", name[0])
    surl = item.xpath('./td[1]/a/@href')
    surlc = url + surl[0]
    html = get_xpath(surlc).xpath('// div[@class="zx-yx-baseinfo"]/a/@href')
    # 讨论区   最多置顶5个  数据有分页（有一部分没有分页）  每页数据为15-20不等 len(zx_html.xpath('//tr[@class="question_cnt_tr"]'))
    zx_url = url + str(html[0])
    zx_html = get_xpath(zx_url)
    # 公告区   没有问题一栏   数据不需要分页
    zx_url1 =url + zx_html.xpath('//div[@class="zx-mid-tabs"]/ul/li/a/@href')[0]
    zx_html1 = get_xpath(zx_url1)
    #精华区    为讨论区置顶内容   数据有分页  每页数据为15 len(zx_html2.xpath('//tr[@class="question_cnt_tr"]'))
    zx_url2 = url + zx_html.xpath('//div[@class="zx-mid-tabs"]/ul/li/a/@href')[1]
    zx_html2 = get_xpath(zx_url2)
    print(name)
    get_ggq(name,zx_html1,wb)
    i = x = 0
    while i == int(x):
        url_test = zx_url2.replace("start-0", "start-" + str(i * 15))
        html_test = get_xpath(url_test)
        if len(html_test.xpath('//tr[@class="question_cnt_tr"]'))==0:
            break
        i += 1
        t = html_test.xpath('//li[@class="lip selected"]//text()')
        x = t[0] * 1
        if i != int(x):
            break
        get_jhq(name,html_test,wb)
    i = x = 0
    zhiding_len = len(zx_html2.xpath('//tr[@class="question_cnt_tr"]'))
    if zhiding_len >= 5:
        zhiding_len = 5
    while i == int(x):
        url_test = zx_url.replace("start-0", "start-" + str(i * 15))
        html_test = get_xpath(url_test)
        if len(html_test.xpath('//tr[@class="question_cnt_tr"]'))==zhiding_len:
            break
        i += 1
        t = html_test.xpath('//li[@class="lip selected"]//text()')
        x = t[0] * 1
        if i != int(x):
            break
        # print(url_test)
        flag = len(zx_html2.xpath('//tr[@class="question_cnt_tr"]'))
        if flag >=5:
            flag = 5
        # time.sleep(1)
        get_tlq(name,flag,html_test,wb)
    # wb.save("资讯信息.xlsx")


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