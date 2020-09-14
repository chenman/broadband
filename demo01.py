import pandas as pd
import numpy as np
import datetime

now = datetime.datetime.now()  # 获取今天零点
today = now - datetime.timedelta(hours=now.hour, minutes=now.minute, seconds=now.second,
                                 microseconds=now.microsecond)  # 获取今日时间
yesterday = now - datetime.timedelta(hours=24, minutes=0, seconds=0, microseconds=0) - datetime.timedelta(
    hours=now.hour, minutes=now.minute, seconds=now.second, microseconds=now.microsecond)  # 获取昨日时间
# print(zeroToday)
# print(yesterday)
df1 = pd.read_excel("cases1.xlsx", sheet_name="Sheet")
df2 = pd.read_excel("cases1.xlsx", sheet_name="sheet1")
dg1 = pd.read_excel("cases1.xlsx", sheet_name="sheet3")

gridmgr = pd.read_excel("cases1.xlsx", sheet_name="gridmgr")

zj = pd.merge(df1, df2, how="left", on="受理部门ID")  # 比对网格长、网格经理（受理清单）
zg = pd.merge(dg1, df2, how="left", on="受理部门ID")  # 比对网格长、网格经理（竣工清单）

# 计算今日受理量
df3 = zj[zj['受理时间'] >= today.strftime("%Y-%m-%d %H:%M:%S")].groupby(zj['网格长']).agg({'工单编号': 'count'}).rename(
    columns={'工单编号': '受理量'})  # 计算网格长当日受理量（受理清单）
df3 = pd.DataFrame(df3)
df3.reset_index(inplace=True)
df3 = pd.merge(gridmgr, df3, on="网格长", how="outer").fillna(0)

df3['受理量'] = df3['受理量'].apply(int)  # 今日受理
print(df3)

# 计算昨日受理量
df4 = zj[(zj['受理时间'] >= yesterday.strftime("%Y-%m-%d %H:%M:%S")) & (
            zj['受理时间'] < today.strftime("%Y-%m-%d %H:%M:%S"))].groupby(zj['网格长']).agg({'工单编号': 'count'}).rename(
    columns={'工单编号': '受理量'})  # 计算网格长当日受理量（受理清单）
df4 = pd.DataFrame(df4)
df4.reset_index(inplace=True)
df4 = pd.merge(gridmgr, df4, on="网格长", how="outer").fillna(0)

df4['受理量'] = df4['受理量'].apply(int)  # 昨日受理
print(df4)


df5 = zg['宽带账号'].groupby(zg['网格长']).count()  # 计算网格长当月竣工量（竣工清单）
df6 = pd.merge(df3, df4, how='outer', on='网格长')  # 合并结果
df7 = pd.merge(df5, df6, how='outer', on='网格长')  # 合并结果
df8 = pd.read_excel("cases1.xlsx", sheet_name="sheet2")
zp = pd.merge(df1, df8, how='left', on='所属网格')  # 按竣工归属网格判断网格长（受理清单）
df9 = zp[zp['受理时间'] >= zeroToday.strftime("%Y-%m-%d %H:%M:%S")]['宽带账号'].groupby(zp['网格长']).count()  # 计算网格当日受理量（受理清单）
df10 = zp[zp['受理时间'] >= yesterday.strftime("%Y-%m-%d %H:%M:%S")]['宽带账号'].groupby(
    zp['网格长']).count() - df9  # 计算网格昨日受理量（受理清单）
zh = pd.merge(dg1, df8, how="left", on="所属网格")  # 按竣工归属网格判断网格长（竣工清单）
df11 = zh['宽带账号'].groupby(zh['网格长']).count()  # 计算网格当月竣工量(竣工清单)
df12 = pd.merge(df9, df10, how='outer', on='网格长')  # 合并结果
df13 = pd.merge(df11, df12, how='outer', on='网格长')  # 合并结果
df14 = pd.merge(df7, df13, how='outer', on='网格长')  # 合并结果
print(df14)
df14.to_excel("cases2.xlsx", index=True)
