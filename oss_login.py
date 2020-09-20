# !usr/bin/env python
# -*- coding:utf-8 -*-
"""
@author:Chen Man
@file:  login.py
@time:  2020/06/17
"""
import urllib.request
import urllib.parse

import requests
import os
import time
import configparser

from aip import AipOcr
import re


# ini文件读取
def read_ini(select_items="baidu-ai"):
    curpath = os.path.dirname(os.path.realpath(__file__))
    cfgpath = os.path.join(curpath, "cfg.ini")

    # 创建管理对象
    conf = configparser.ConfigParser()
    # 读ini文件
    conf.read(cfgpath, encoding="utf-8-sig")  # python3
    try:
        items = conf.items(select_items)
        return (items)
    except Exception:
        return -1


def header_format(raw_headers):
    headers = dict([line.split(": ", 1) for line in raw_headers.strip().split("\n")])
    # print(headers)
    return headers


# 通过百度api识别验证码，并放回结果
def verification_code(filePath):
    baidu_key = read_ini("baidu-ai")
    APP_ID = baidu_key[0][1]
    API_KEY = baidu_key[1][1]
    SECRET_KEY = baidu_key[2][1]
    client = AipOcr(APP_ID, API_KEY, SECRET_KEY)
    with open(filePath, 'rb') as fp:
        image = fp.read()
        a = client.basicGeneral(image);
        try:
            code = a['words_result'][0]['words']
            # 只保留字母和数字
            c = filter(str.isalnum, code)
            code = ''.join(list(c))
            print("验证码：" + code)
            return (code)
        except Exception:
            print("请检查你的百度AI相应key信息是否正确？")


def login_try(username, password, cookies):
    code_header = '''Accept: */*
Referer: http://10.53.160.88:20040/oss-web/index.jsp
Accept-Language: zh-CN
User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)
Accept-Encoding: gzip, deflate
Host: 10.53.160.88:20040
Connection: Keep-Alive'''

    code_url = 'http://10.53.160.88:20040/oss-web/IOM-OAAS-SERVICE/oaas/code.do?time=%s' % (int(time.time() * 1000))
    response = requests.get(code_url, headers=header_format(code_header), cookies=cookies)
    # cookies = response.cookies

    img_path = os.getcwd() + "\code.jpg"
    with open(img_path, "wb") as img:
        img.write(response.content)
        img.close()
    code = verification_code(img_path)
    # print(code)

    login_header = '''Content-Type: application/x-www-form-urlencoded; charset=UTF-8
Accept: */*
X-Requested-With: XMLHttpRequest
Referer: http://10.53.160.88:20040/oss-web/index.jsp
Accept-Language: zh-cn
Accept-Encoding: gzip, deflate
User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)
Host: 10.53.160.88:20040
Connection: Keep-Alive
Cache-Control: no-cache'''

    login_url = 'http://10.53.160.88:20040/oss-web/IOM-OAAS-SERVICE/oaas/login'
    # post_data = {'user': username, 'pw': password, 'certCode': code}
    post_data = 'user=' + username + '&pw=' + password + '&certCode=' + code
    try:
        response = requests.post(login_url, data=post_data, cookies=cookies, headers=header_format(login_header))
        # cookies = response.cookies
        # print(response.text)
        if response.status_code != 200:
            time.sleep(3)
            return login_try(username, password, cookies)
        else:
            staff_id = response.json()['user']['staffId']
            return staff_id
    except requests.HTTPError as e:
        print(e.reason)


def login(username, password):
    index_header = '''Accept: */*
Accept-Encoding: gzip, deflate
Accept-Language: zh-CN
Connection: Keep-Alive
Host: 10.53.160.88:20040
User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)
'''
    index_url = 'http://10.53.160.88:20040/oss-web/index.jsp'

    response = requests.get(index_url, headers=header_format(index_header))

    cookies = response.cookies
    cookies_dict = requests.utils.dict_from_cookiejar(cookies)
    token = cookies_dict["ZXRC-OSS-SESSION"]
    login_try(username, password, cookies)

    refer_url = 'http://10.53.160.88:20040/resweb/sso.do?token={0}&page=/component/dynamicPage/dynamicQuery/query.jsp?queryId=11114&roleId=4018&menuInterface=WEB_MAIN&preConditions=[]&resTypeId=2530&needMenu=true&specialityId=90&multiOption=true&isPopMenu=true&regionAndSpecXML='.format(
        token)
    refer_header = '''Accept: */*
Referer: http://10.53.160.88:20040/oss-web/main.jsp
Accept-Language: zh-CN
User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)
Accept-Encoding: gzip, deflate
Host: 10.53.160.88:20040
Connection: Keep-Alive'''
    s = requests.Session()
    response = s.get(refer_url, headers=header_format(refer_header), cookies=cookies, allow_redirects=False)
    cookies = dict(response.cookies, **cookies)
    location = response.headers['Location']
    # print(cookies)
    # print(location)
    # print(response.text)

    # refer_url2 = urllib.parse.unquote(location)
    refer_url2 = location
    # print(refer_url2)
    refer_header2 = '''Accept: */*
Referer: http://10.53.160.88:20040/oss-web/main.jsp
Accept-Language: zh-CN
User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)
Accept-Encoding: gzip, deflate
Connection: Keep-Alive
Host: 10.53.160.88:8188'''
    response = s.get(refer_url2, headers=header_format(refer_header2), cookies=cookies)
    # print(response.text)

    sso_login_url = re.findall(r'<form action="(.*)" method="post"', response.text)[0]
    sso_login_header = '''Accept: */*
Referer: {0}
Accept-Language: zh-CN
User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)
Content-Type: application/x-www-form-urlencoded
Accept-Encoding: gzip, deflate
Host: 10.53.160.88:8188
Content-Length: 0
Connection: Keep-Alive'''.format(refer_url2)
    response = requests.post(sso_login_url, headers=header_format(sso_login_header), cookies=cookies)
    return cookies


def export_data(request_xml, cookies, filename):
    export_url = 'http://10.53.160.88:20040/IOMPROJ/FaultExportForwardServlet'
    export_header = '''Accept: image/gif, image/jpeg, image/pjpeg, application/x-ms-application, application/xaml+xml, application/x-ms-xbap, */*
Accept-Encoding: gzip, deflate
Accept-Language: zh-CN
Cache-Control: no-cache
Connection: Keep-Alive
Content-Type: application/x-www-form-urlencoded
Host: 10.53.160.88:20040
Referer: http://10.53.160.88:20040/IOMPROJ/order/orderNotRealTimeSearch.jsp
User-Agent: Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 10.0; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729)'''

    post_data = {'forwardServlet': 1, 'functionXml': request_xml}
    response = requests.post(export_url, headers=header_format(export_header), cookies=cookies, data=post_data)

    file = open(filename, 'wb')
    file.write(response.content)
    file.close()


if __name__ == '__main__':
    pass
