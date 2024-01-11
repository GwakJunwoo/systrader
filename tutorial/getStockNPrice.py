import win32com.client
from checkConnect import *

def getStockNPrice(code, N):

    def requestData(obj):
            rqStatus = obj.GetDibStatus()
            if rqStatus != 0: return False

            stockPrice = []
            count = obj.GetHeaderValue(1)
            for i in range(count):
                 date = obj.GetDataValue(0, i)
                 open = obj.GetDataValue(1, i)
                 high = obj.GetDataValue(2, i)
                 low = obj.GetDataValue(3, i)
                 close = obj.GetDataValue(4, i)
                 diff = obj.GetDataValue(5, i)
                 vol = obj.GetDataValue(6, i)
                 
                 stockPrice.append((date, open, high, low, close, diff, vol))

            return stockPrice

    isConnect = checkConnect()
    if isConnect == False: exit()

    obj = win32com.client.Dispatch("DsCbo1.StockWeek")
    obj.SetInputValue(0, code)
    obj.BlockRequest()

    # 통신상태 확인 및 에러처리
    rqStatus = obj.GetDibStatus()
    if rqStatus != 0: exit()

    # 최초 데이터 요청
    stockPrice = requestData(obj)
    if stockPrice == False: exit()

    nextCount = 1
    while obj.Continue:
        nextCount += 1
        if nextCount > N: break
        stockPrice = stockPrice + requestData(obj)
        if stockPrice == False: exit()

    return stockPrice

