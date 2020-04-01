# -*- coding:utf-8 -*-
import re
import os
import requests
from lxml import html
from urllib import parse
import csv
import xlrd
from copy import copy

# 搜索关键字，这里只爬取了数据挖掘的数据，读者可以更换关键字爬取其他行业数据
city = {"北京": '010000',
        "上海": '020000',
        "广州": '030200',
        "深圳": '040000',
        "成都": '090200', "杭州": '080200',"南昌": '130200' }

#北京 010000
#上海 020000
#广州 030200
#深圳 040000
#成都 090200
#杭州 080200
#南昌 130200
# 伪装爬取头部，以防止被网站禁止
headers = {'Host': 'search.51job.com',
           'Upgrade-Insecure-Requests': '1',
           'User-Agent': ['Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko)\
         Chrome/63.0.3239.132 Safari/537.36',
       "Mozilla/5.0 (compatible; Baiduspider/2.0; +http://www.baidu.com/search/spider.html)",
       "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1; AcooBrowser; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
       "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.0; Acoo Browser; SLCC1; .NET CLR 2.0.50727; Media Center PC 5.0; .NET CLR 3.0.04506)",
       "Mozilla/4.0 (compatible; MSIE 7.0; AOL 9.5; AOLBuild 4337.35; Windows NT 5.1; .NET CLR 1.1.4322; .NET CLR 2.0.50727)",
       "Mozilla/5.0 (Windows; U; MSIE 9.0; Windows NT 9.0; en-US)",
       "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Win64; x64; Trident/5.0; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 2.0.50727; Media Center PC 6.0)",
       "Mozilla/5.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0; WOW64; Trident/4.0; SLCC2; .NET CLR 2.0.50727; .NET CLR 3.5.30729; .NET CLR 3.0.30729; .NET CLR 1.0.3705; .NET CLR 1.1.4322)",
       "Mozilla/4.0 (compatible; MSIE 7.0b; Windows NT 5.2; .NET CLR 1.1.4322; .NET CLR 2.0.50727; InfoPath.2; .NET CLR 3.0.04506.30)",
       "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN) AppleWebKit/523.15 (KHTML, like Gecko, Safari/419.3) Arora/0.3 (Change: 287 c9dfb30)",
       "Mozilla/5.0 (X11; U; Linux; en-US) AppleWebKit/527+ (KHTML, like Gecko, Safari/419.3) Arora/0.6",
       "Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.8.1.2pre) Gecko/20070215 K-Ninja/2.1.1",
       "Mozilla/5.0 (Windows; U; Windows NT 5.1; zh-CN; rv:1.9) Gecko/20080705 Firefox/3.0 Kapiko/3.0",
       "Mozilla/5.0 (X11; Linux i686; U;) Gecko/20070322 Kazehakase/0.4.5",
       "Mozilla/5.0 (X11; U; Linux i686; en-US; rv:1.9.0.8) Gecko Fedora/1.9.0.8-1.fc10 Kazehakase/0.5.6",
       "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/535.11 (KHTML, like Gecko) Chrome/17.0.963.56 Safari/535.11",
       "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_7_3) AppleWebKit/535.20 (KHTML, like Gecko) Chrome/19.0.1036.7 Safari/535.20",
       "Opera/9.80 (Macintosh; Intel Mac OS X 10.6.8; U; fr) Presto/2.9.168 Version/11.52"
]}


# 获取职位详细页面
def get_links(page, city_code, key):
    url = 'http://search.51job.com/list/'+ str(city_code) +',000000,0000,00,9,99,' + key + ',2,' + str(page) + '.html'
    r = requests.get(url, headers, timeout=10)
    s = requests.session()
    s.keep_alive = False
    r.encoding = 'gbk'
    reg = re.compile(r'class="t1 ">.*? <a target="_blank" title=".*?" href="(.*?)".*? <span class="t2">', re.S)
    links = re.findall(reg, r.text)
    return links


# 多页处理，下载到文件
def get_content(link, writer):
    r1 = requests.get(link, headers, timeout=10)
    s = requests.session()
    s.keep_alive = False
    r1.encoding = 'gb2312'
    t1 = html.fromstring(r1.text)
    # print(link)
    job = t1.xpath('//div[@class="tHeader tHjob"]//h1/text()')[0].strip()
    print(job)
    company = t1.xpath('//p[@class="cname"]/a/text()')[0].strip()
    # print(company)
    salary = t1.xpath('//div[@class="cn"]//strong/text()')[0].strip()
    # print(salary)
    area = t1.xpath('//p[@class="msg ltype"]/text()')[0].strip()
    # print(area)
    experience = t1.xpath('//p[@class="msg ltype"]/text()')[1].strip()
    # print(experience)
    education = t1.xpath('//p[@class="msg ltype"]/text()')[2].strip()
    # print(education)
    company_type = t1.xpath('//p[@class="at"]/text()')[0].strip()
    # print(company_type)
    company_size = t1.xpath('//p[@class="at"]/text()')[1].strip()
    # print(companyscale)
    direction = t1.xpath('//div[@class="com_tag"]/p/a/text()')[0].strip()
    # print(direction)
    address = t1.xpath('//div[@class="bmsg inbox"]//text()')[2].strip()
    describe = t1.xpath('//div[@class="bmsg job_msg inbox"]//text()')
    # print(describe)
    writer.writerow((link, job, company, salary, area, experience, education, company_type, company_size,
                     direction, address, describe))
    return True


def write_excel_xls_append(items, file_name, path):
    index = len(items)  # 获取需要写入数据的行数
    workbook = xlrd.open_workbook(path + file_name)  # 打开工作簿
    sheets = workbook.sheet_names()  # 获取工作簿中的所有表格
    worksheet = workbook.sheet_by_name(sheets[0])  # 获取工作簿中所有表格中的的第一个表格
    rows_old = worksheet.nrows  # 获取表格中已存在的数据的行数
    new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格
    for i in range(0, index):
        item = items[i]
        # 追加写入数据，注意是从i+rows_old行开始写入
        new_worksheet.write(i + rows_old, 0, item['link'])
        new_worksheet.write(i + rows_old, 1, item['job'])
        new_worksheet.write(i + rows_old, 2, item['company'])
        new_worksheet.write(i + rows_old, 3, item['salary'])
        new_worksheet.write(i + rows_old, 4, item['area'])
        new_worksheet.write(i + rows_old, 5, item['experience'])
        new_worksheet.write(i + rows_old, 6, item['education'])
        new_worksheet.write(i + rows_old, 7, item['companyType'])
        new_worksheet.write(i + rows_old, 8, item['companySize'])
        new_worksheet.write(i + rows_old, 9, item['direction'])
        new_worksheet.write(i + rows_old, 10, item['address'])
        new_worksheet.write(i + rows_old, 11, item['describe'])

    new_workbook.save(path + file_name)  # 保存工作簿
    print("追加数据成功！")


def main():
    # 主调动函数
    # 爬取前三页信息
    cityName = str(input('请输入查找的地区：'))
    keyWord = str(input('请输入查找的职位关键字：'))
    needPage = int(input('请输入要爬取的页数(页/50条)：'))
    # 编码调整，如将“数据挖掘”编码成%25E6%2595%25B0%25E6%258D%25AE%25E6%258C%2596%25E6%258E%2598
    key = parse.quote(parse.quote(keyWord))

    # .csv文件，进行写入操作
    file_name = cityName+'-'+keyWord+'.csv'
    path = './前程无忧/'
    if not os.path.exists(path):
        os.mkdir(path)
    csvFile = open(path + file_name, 'w', newline='')
    writer = csv.writer(csvFile)
    writer.writerow(('link', 'job', 'company', 'salary', 'area', 'experience', \
                     'education', 'companyType', 'companySize', 'direction', 'address', 'describe'))

    for i in range(1, needPage + 1):
        print('正在爬取第{}页信息'.format(i))
        city_code = city.get(cityName, '000000')
        links = get_links(i + 1, city_code, key)
        for link in links:
            try:
                get_content(link, writer)
            except:
                print("数据有缺失值")
                continue
    # 关闭写入文件
    csvFile.close()


if __name__ == '__main__':
    main()
