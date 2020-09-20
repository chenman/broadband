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
import datetime
import pandas as pd

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
    # print(int(time.time() * 1000))
    cookies = login('chenman', '!EIje*62')
    now = datetime.datetime.now()
    init_date = datetime.date.today()
    pre_date = init_date - datetime.timedelta(days=1)
    current_month_first_day = datetime.date(year=now.year, month=now.month, day=1)
    # print(pre_date.strftime('%Y-%m-%d %H:%M:%S'))
    # print(now.strftime('%Y-%m-%d %H:%M:%S'))
    # 2020-09-10 00:00:00

    accept_request_xml = '<?xml version="1.0" encoding="gb2312"?><requestdata><parameter ' \
                         'name="reportCode">orderNotRealTime</parameter><parameter ' \
                         'name="ttOrgId">10000000</parameter><parameter name="priv">infoHide</parameter><parameter ' \
                         'name="areaIds">4,102,41,42,43,44,45</parameter><parameter ' \
                         'name="serviceIds">220323,220372</parameter><parameter ' \
                         'name="searchType">OrdMoni</parameter><parameter ' \
                         'name="startDate">%s</parameter><parameter name="endDate">%s</parameter><parameter ' \
                         'name="DateType">1</parameter><parameter ' \
                         'name="searchFlag">true</parameter><parameter name="staffId">223214</parameter><parameter ' \
                         'name="pageIndex">1</parameter><parameter name="pageSize">30000</parameter></requestdata> ' \
                         % (current_month_first_day.strftime('%Y-%m-%d %H:%M:%S'), now.strftime('%Y-%m-%d %H:%M:%S'))

    accept_list_file = now.strftime("%Y%m%d%H%M%S-accept") + '.xlsx'
    export_data(accept_request_xml, cookies, filename=accept_list_file)

    finish_request_xml = '<?xml version="1.0" encoding="gb2312"?><requestdata><parameter ' \
                         'name="reportCode">orderNotRealTime</parameter><parameter ' \
                         'name="ttOrgId">10000000</parameter><parameter name="priv">infoHide</parameter><parameter ' \
                         'name="areaIds">4,102,41,42,43,44,45</parameter><parameter ' \
                         'name="serviceIds">220323</parameter><parameter name="omStates">10F</parameter><parameter ' \
                         'name="searchType">OrdMoni</parameter><parameter name="startDate">%s</parameter><parameter ' \
                         'name="endDate">%s</parameter><parameter ' \
                         'name="DateType">2</parameter><parameter name="searchFlag">true</parameter><parameter ' \
                         'name="staffId">223214</parameter><parameter name="pageIndex">1</parameter><parameter ' \
                         'name="pageSize">30000</parameter></requestdata>' \
                         % (current_month_first_day.strftime('%Y-%m-%d %H:%M:%S'), now.strftime('%Y-%m-%d %H:%M:%S'))

    finish_list_file = now.strftime("%Y%m%d%H%M%S-finish") + '.xlsx'
    export_data(finish_request_xml, cookies, filename=finish_list_file)

    cfg_org = pd.read_excel("config.xlsx", sheet_name="sheet1")
    cfg_grd = pd.read_excel("config.xlsx", sheet_name="sheet2")
    cfg_mgr = pd.read_excel("config.xlsx", sheet_name="sheet3")

    # print(cfg_org)
    # print(cfg_grd)
    # print(cfg_mgr)

    accept_list = pd.read_excel(accept_list_file, '全量工单信息')

    accept_list_2d = accept_list[accept_list['受理时间'] >= pre_date.strftime("%Y-%m-%d %H:%M:%S")]

    al = pd.merge(accept_list_2d, cfg_org, how="left", on="受理部门ID")  # 比对网格长、网格经理（受理清单）

    # 计算今日受理量
    day_accept = al[al['受理时间'] >= init_date.strftime("%Y-%m-%d %H:%M:%S")].groupby(al['网格长']).agg(
        {'工单编号': 'count'}).rename(
        columns={'工单编号': '网点当日受理量'})  # 计算网格长当日受理量（受理清单）
    day_accept = pd.DataFrame(day_accept)
    day_accept.reset_index(inplace=True)
    day_accept = pd.merge(cfg_mgr, day_accept, on="网格长", how="outer").fillna(0)
    day_accept['网点当日受理量'] = day_accept['网点当日受理量'].apply(int)  # 今日受理量
    # print(day_accept)

    # 计算昨日受理量
    pre_accept = al[(al['受理时间'] >= pre_date.strftime("%Y-%m-%d %H:%M:%S")) & (
            al['受理时间'] < init_date.strftime("%Y-%m-%d %H:%M:%S"))].groupby(al['网格长']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '网点昨日受理量'})  # 计算网格长当日受理量（受理清单）
    pre_accept = pd.DataFrame(pre_accept)
    pre_accept.reset_index(inplace=True)
    pre_accept = pd.merge(cfg_mgr, pre_accept, on="网格长", how="outer").fillna(0)
    pre_accept['网点昨日受理量'] = pre_accept['网点昨日受理量'].apply(int)  # 昨日受理量
    # print(pre_accept)

    # 竣工信息
    finish_list = pd.read_excel(finish_list_file, '全量工单信息')

    # 网点竣工信息
    ofl = pd.merge(finish_list, cfg_org, how="left", on="受理部门ID")  # 比对网格长、网格经理（受理清单）

    # 网点当月竣工量
    o_month_finish = ofl.groupby(ofl['网格长']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '网点当月竣工量'})
    o_month_finish = pd.DataFrame(o_month_finish)
    o_month_finish.reset_index(inplace=True)
    o_month_finish = pd.merge(cfg_mgr, o_month_finish, on="网格长", how="outer").fillna(0)
    o_month_finish['网点当月竣工量'] = o_month_finish['网点当月竣工量'].apply(int)  # 昨日受理量
    # print(o_month_finish)

    # 网格当月竣工量
    gfl = pd.merge(finish_list, cfg_grd, how="left", on="所属网格")

    # 网格当月竣工量
    g_month_finish = gfl.groupby(gfl['网格长']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '网格当月竣工量'})
    g_month_finish = pd.DataFrame(g_month_finish)
    g_month_finish.reset_index(inplace=True)
    g_month_finish = pd.merge(cfg_mgr, g_month_finish, on="网格长", how="outer").fillna(0)
    g_month_finish['网格当月竣工量'] = g_month_finish['网格当月竣工量'].apply(int)  # 昨日受理量
    # print(g_month_finish)

    # 网格受理信息
    gal = pd.merge(accept_list_2d, cfg_grd, how="left", on="所属网格")  # 比对网格长、网格经理（受理清单）

    # 计算网格今日受理量
    g_day_accept = gal[gal['受理时间'] >= init_date.strftime("%Y-%m-%d %H:%M:%S")].groupby(gal['网格长']).agg(
        {'工单编号': 'count'}).rename(
        columns={'工单编号': '网格当日受理量'})  # 计算网格长当日受理量（受理清单）
    g_day_accept = pd.DataFrame(g_day_accept)
    g_day_accept.reset_index(inplace=True)
    g_day_accept = pd.merge(cfg_mgr, g_day_accept, on="网格长", how="outer").fillna(0)
    g_day_accept['网格当日受理量'] = g_day_accept['网格当日受理量'].apply(int)  # 今日受理量
    # print(g_day_accept)

    # 计算网格昨日受理量
    g_pre_accept = gal[(gal['受理时间'] >= pre_date.strftime("%Y-%m-%d %H:%M:%S")) & (
            gal['受理时间'] < init_date.strftime("%Y-%m-%d %H:%M:%S"))].groupby(gal['网格长']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '网格昨日受理量'})  # 计算网格长当日受理量（受理清单）
    g_pre_accept = pd.DataFrame(g_pre_accept)
    g_pre_accept.reset_index(inplace=True)
    g_pre_accept = pd.merge(cfg_mgr, g_pre_accept, on="网格长", how="outer").fillna(0)
    g_pre_accept['网格昨日受理量'] = g_pre_accept['网格昨日受理量'].apply(int)  # 昨日受理量
    # print(g_pre_accept)

    result = pd.merge(o_month_finish, day_accept, on='网格长', how='left')
    result = pd.merge(result, pre_accept,  on='网格长', how='left')
    result = pd.merge(result, g_month_finish,  on='网格长', how='left')
    result = pd.merge(result, g_day_accept,  on='网格长', how='left')
    result = pd.merge(result, g_pre_accept,  on='网格长', how='left')
    # 保存网格报表
    result.to_excel("grid_static.xlsx", index=True, sheet_name='网格')

    # 区县报表
    month_accept = accept_list.groupby(accept_list['区县']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '区县当月受理量'})  # 区县月受理量
    month_accept = pd.DataFrame(month_accept)
    month_accept.reset_index(inplace=True)

    today_accept = accept_list[accept_list['受理时间'] >= init_date.strftime("%Y-%m-%d %H:%M:%S")]\
        .groupby(accept_list['区县']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '区县当日受理'})
    today_accept = pd.DataFrame(today_accept)
    today_accept.reset_index(inplace=True)

    month_finish = finish_list.groupby(finish_list['区县']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '区县当月竣工量'})  # 区县月竣工量
    month_finish = pd.DataFrame(month_finish)
    month_finish.reset_index(inplace=True)

    today_finish = finish_list[finish_list['竣工时间'] >= init_date.strftime("%Y-%m-%d %H:%M:%S")]\
        .groupby(finish_list['区县']).agg({'工单编号': 'count'}).rename(
        columns={'工单编号': '区县当日竣工'})
    today_finish = pd.DataFrame(today_finish)
    today_finish.reset_index(inplace=True)

    result = pd.merge(month_accept, today_accept, on='区县', how='left')
    result = pd.merge(result, month_finish, on='区县', how='left')
    result = pd.merge(result, today_finish, on='区县', how='left')
    result.to_excel("county_static.xlsx", index=True, sheet_name='区县')
