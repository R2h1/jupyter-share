import requests
from bs4 import BeautifulSoup       #网页解析获取数据
import re
import time
import xlwt
import math

sleepTime = 2

#爬取腾讯视频电影数据
def get_ten(count= 1000, pageSize = 25):
    url = "https://v.qq.com/channel/movie?listpage=1&channel=movie&sort=18&_all=1"        #链接
    param = {                                                                             #参数字典
        'offset': 0,
        'pagesize': pageSize
    }
    headers={                                                                            #UA伪装
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '+
                       'AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.85 Safari/537.36'
    }
    offset = 0                                                                           #拼接url
    dataRes = []
    findLink = re.compile(r'href="(.*?)"')  # 链接
    findName = re.compile(r'title="(.*?)"')  # 影片名
    for i in range(0, math.ceil(count / pageSize)):
        print("\r开始获取第" + str(offset + 1) + "-" +  str(offset + pageSize) + "部电影")
        res = requests.get(url=url, params=param , headers=headers)       #编辑request请求
        res.encoding='utf-8'                                           #设置返回数据的编码格式为utf-8
        html = BeautifulSoup(res.text,"html.parser")      #BeautifulSoup解析
        part_html = html.find_all(r"a", class_="figure")               #找到整个html界面里a标签对应的html代码，返回值是一个list
        if (len(part_html) == 0):
            print("页面返回空！已获取条数：" + str(len(dataRes)))
            return dataRes
        offset = offset + pageSize                                         #修改参数字典+25部电影
        param['offset'] = offset
        for i in part_html:                                            #遍历每一个part_html
            words = str(i)
            name = re.findall(findName, words)# 添加影片名
            # score = re.findall(findScore, words)# 添加评分
            link = re.findall(findLink, words)# 添加链接
            list_ = []
            # if(len(score)==0):
            #     score.insert(0,"暂无评分")
            for i in dataRes:
                if name[0] in i[0]:
                    name.insert(0,name[0]+"（其他版本）")
            list_.append(name[0])
            # list_.append(score[0])
            list_.append(link[0])
            dataRes.append(list_)
        time.sleep(sleepTime)
    # print(dataRes)      #打印最终结果
    return dataRes


#  将数据写入新文件
def data_write(file_path, datas, count):
    if (len(datas) < count):
        print("已获取条数：" + str(len(datas)) + ",请重试！")
        return
    f = xlwt.Workbook()
    sheet1 = f.add_sheet(u'sheet1',cell_overwrite_ok=True) #创建sheet
    
    #将数据写入第 i 行，第 j 列
    i = 0
    for data in datas:
        for j in range(len(data)):
            sheet1.write(i, j, data[j])
        i = i + 1

    f.save(file_path) #保存文件
    print("写入 " + file_path[2:] + " 成功")

if __name__ == '__main__':
    count = 1000 # 总数
    pageSize = 30 # 每页数量 不要超过30, 否则和页面表现不一致
    res = get_ten(count, pageSize)
    file_path = './腾讯视频电影片库最近热播' + 'Top' + str(count) + '.xls'
    data_write(file_path, res, count)
