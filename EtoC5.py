import xlrd

from translate import Translator

import re

import json

import requests

import xlwt
 
def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook('C:/Users/ZERONOW/Desktop/engTOch/file.xls')

    sheet1 = workbook.sheet_by_index(0)  # sheet索引从0开始

    f = xlwt.Workbook() #创建待写入工作薄
    sheet2 = f.add_sheet(u'ONE',cell_overwrite_ok=True) #创建待写入sheet

    print(">>>>>>>>>>>----------翻译开始----------<<<<<<<<<<")

    for i in range(0,sheet1.nrows):
        rows = sheet1.row_values(i)

        #全部转换成str格式
        rows = [str(ele) for ele in rows]

        print("原文：",rows)

        listCH = []

        #拆解到每一个框中的元素再进行翻译
        for ele in rows:

            #只翻译英文单词
            flag = re.fullmatch(r"[A-Za-z][A-Za-z]*",ele)

            if ele.strip()!='':#如果ele不为空
                if flag!=1:
                    list_trans = translate(ele)

                    #print(list_trans)

                    result = json.loads(list_trans)
                    # print(result['translateResult'][0][0]['tgt'])

                    listCH.append(result['translateResult'][0][0]['tgt'])
                else:
                    # print(ele)
                    listCH.append(ele)
            else:
                # print(ele)
                listCH.append(ele)

        #打印出list方便输入到excel
        print("译文：",listCH)

        #写入excel
        for ii in range(len(listCH)):
            sheet2.write(i,ii,listCH[ii]) 
        f.save('C:/Users/ZERONOW/Desktop/engTOch/NEW.xls')#保存文件
    print("翻译后的Excel文件保存在“C:/Users/ZERONOW/Desktop/engTOch/NEW.xls”")
    print(">>>>>>>>>>----------翻译结束----------<<<<<<<<<<")

# 翻译函数，word 需要翻译的内容
def translate(word):
    # 有道词典 api
    url = 'http://fanyi.youdao.com/translate?smartresult=dict&smartresult=rule&smartresult=ugc&sessionFrom=null'
    # 传输的参数，其中 i 为需要翻译的内容
    key = {
        'type': "AUTO",
        'i': word,
        "doctype": "json",
        "version": "2.1",
        "keyfrom": "fanyi.web",
        "ue": "UTF-8",
        "action": "FY_BY_CLICKBUTTON",
        "typoResult": "true"
    }
    # key 这个字典为发送给有道词典服务器的内容
    response = requests.post(url, data=key)
    # 判断服务器是否相应成功
    if response.status_code == 200:
        # 然后相应的结果
        return response.text
    else:
        print("有道词典调用失败")
        # 相应失败就返回空
        return None

if __name__ == '__main__':
    read_excel()