import datetime
import pandas as pd

now = datetime.datetime.now()  # 获取今天零点
zeroToday = now - datetime.timedelta(hours=now.hour, minutes=now.minute, seconds=now.second,
                                     microseconds=now.microsecond)  # 获取今日时间
yesterday = now - datetime.timedelta(hours=24, minutes=0, seconds=0, microseconds=0) - datetime.timedelta(
    hours=now.hour, minutes=now.minute, seconds=now.second, microseconds=now.microsecond)  # 获取昨日时间
print(zeroToday)
print(yesterday)
df1 = pd.read_excel("cases1.xlsx", sheet_name="Sheet")
df2 = pd.read_excel("cases1.xlsx", sheet_name="sheet1")
dg1 = pd.read_excel("cases1.xlsx", sheet_name="sheet3")
zj = pd.merge(df1, df2, how="left", on="受理部门ID")  # 比对网格长、网格经理、网点名称（受理清单）
zg = pd.merge(dg1, df2, how="left", on="受理部门ID")  # 比对网格长、网格经理、网点名称（竣工清单）
df3 = zj[zj['受理时间'] >= zeroToday.strftime("%Y-%m-%d %H:%M:%S")]['宽带账号'].groupby(zj['渠道名称']).count()  # 计算网点当日受理量（受理清单）
df4 = zj[zj['受理时间'] >= yesterday.strftime("%Y-%m-%d %H:%M:%S")]['宽带账号'].groupby(
    zj['渠道名称']).count() - df3  # 计算网点昨日受理量（受理清单）
df5 = zg['宽带账号'].groupby(zg['渠道名称']).count()  # 计算网点当月竣工量（竣工清单）
df6 = pd.merge(df3, df4, how='outer', on='渠道名称')  # 合并结果
df7 = pd.merge(df5, df6, how='outer', on='渠道名称')  # 合并结果
df8 = pd.merge(df2, df7, how='outer', on='渠道名称')  # 合并结果
df8.to_excel("case5.xlsx", index=True)
