


import ExcelUtil
from ExcelUtil import show
from ExcelUtil import filterbycontains
from ExcelUtil import filterbynocontains
from datetime import datetime


FILE_EXCEL_0 = '装维管控日报模版.xlsx'
FILE_EXCEL_1 = '表1装移机工单一览表.xlsx'
FILE_EXCEL_2 = '外线工单监控箱'
FILE_EXCEL = ''


def isindatetime(dt,start=None,stop=None):
    flag = True
    if start is not None:
        if (dt - start).days < 0:
            flag = False
    if stop is not None:
        if (stop - dt).days < 0:
            flag = False
    return flag

# 筛选并返回在时间范围内的数据
def filterbyindatetime(data=[],idx=0,start=None,stop=None):
    result = []
    for row in data:
        if isindatetime(row[idx],start=start,stop=stop): #在时间之内
            result.append(row)
        else :
            pass
    return result

def main():
    eu1 = ExcelUtil.ExcelUtil(filename=FILE_EXCEL_1,new=False)
    data = eu1.getcolsbyname(items=['机楼','局向','业务类型','施工类型','归档时间']) # 筛选出含items
    result1 = filterbycontains(data = data, items=['新装']) # 筛选出含'新装'
    result2 = filterbynocontains(data = result1, items=['普通电话新装']) # 筛选出不含'普通电话新装'
    result3 = filterbyindatetime(data = result2, idx = 4 , start = datetime(2018, 3, 1) , stop = datetime(2018, 3, 2)) # 筛选时间内的数据

    # a1 = filterbycontains(data = data, items=['ADSL']) # 筛选出含'新装'
    show(result3)
    eu1.close()
    pass

if __name__=='__main__':
    # a = datetime(2018, 3, 4)
    # b = datetime(2018, 3, 5)
    # print((b-a).days)
    # start = datetime(2018, 3, 1)
    # stop = datetime(2018, 3, 8)
    # a = datetime(2018,3,5)
    # print(isindatetime(a,start=start,stop=stop))
    # a = datetime.now()
    # print(a)
    # b = datetime(2018,3,9)
    # print(b)
    # c = a - b
    # print(c.days)
    # exit(0)
    print('main')
    main()
