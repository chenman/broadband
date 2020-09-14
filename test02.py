
import datetime
import urllib.parse

if __name__ == '__main__':
    # now = datetime.datetime.now()
    # current_month_first_day = datetime.date(year=now.year, month=now.month, day=1)
    # print(current_month_first_day.strftime("%Y%m%d%H%M%S"))
    # print(urllib.parse.unquote_plus(
    #    '%3C%3Fxml+version%3D%221.0%22+encoding%3D%22gb2312%22%3F%3E%3Crequestdata%3E%3Cparameter+name%3D%22reportCode%22%3EorderNotRealTime%3C%2Fparameter%3E%3Cparameter+name%3D%22ttOrgId%22%3E10000000%3C%2Fparameter%3E%3Cparameter+name%3D%22priv%22%3EinfoHide%3C%2Fparameter%3E%3Cparameter+name%3D%22areaIds%22%3E4%2C102%2C41%2C42%2C43%2C44%2C45%3C%2Fparameter%3E%3Cparameter+name%3D%22serviceIds%22%3E220323%3C%2Fparameter%3E%3Cparameter+name%3D%22omStates%22%3E10F%3C%2Fparameter%3E%3Cparameter+name%3D%22searchType%22%3EOrdMoni%3C%2Fparameter%3E%3Cparameter+name%3D%22startDate%22%3E2020-09-01+00%3A00%3A00%3C%2Fparameter%3E%3Cparameter+name%3D%22endDate%22%3E2020-09-14+21%3A56%3A02%3C%2Fparameter%3E%3Cparameter+name%3D%22DateType%22%3E2%3C%2Fparameter%3E%3Cparameter+name%3D%22searchFlag%22%3Etrue%3C%2Fparameter%3E%3Cparameter+name%3D%22staffId%22%3E223214%3C%2Fparameter%3E%3Cparameter+name%3D%22pageIndex%22%3E1%3C%2Fparameter%3E%3Cparameter+name%3D%22pageSize%22%3E10000%3C%2Fparameter%3E%3C%2Frequestdata%3E'))

    print(urllib.parse.unquote_plus('%3C%3Fxml+version%3D%221.0%22+encoding%3D%22gb2312%22%3F%3E%3Crequestdata%3E%3Cparameter+name%3D%22reportCode%22%3EorderNotRealTime%3C%2Fparameter%3E%3Cparameter+name%3D%22ttOrgId%22%3E10000000%3C%2Fparameter%3E%3Cparameter+name%3D%22priv%22%3EinfoHide%3C%2Fparameter%3E%3Cparameter+name%3D%22areaIds%22%3E4%2C102%2C41%2C42%2C43%2C44%2C45%3C%2Fparameter%3E%3Cparameter+name%3D%22serviceIds%22%3E220323%2C220372%3C%2Fparameter%3E%3Cparameter+name%3D%22searchType%22%3EOrdMoni%3C%2Fparameter%3E%3Cparameter+name%3D%22startDate%22%3E2020-09-14+00%3A00%3A00%3C%2Fparameter%3E%3Cparameter+name%3D%22endDate%22%3E2020-09-15+14%3A13%3A59%3C%2Fparameter%3E%3Cparameter+name%3D%22DateType%22%3E1%3C%2Fparameter%3E%3Cparameter+name%3D%22searchFlag%22%3Etrue%3C%2Fparameter%3E%3Cparameter+name%3D%22staffId%22%3E223214%3C%2Fparameter%3E%3Cparameter+name%3D%22pageIndex%22%3E1%3C%2Fparameter%3E%3Cparameter+name%3D%22pageSize%22%3E10000%3C%2Fparameter%3E%3C%2Frequestdata%3E'))
