import win32com.client
from checkConnect import *

def getStockList():
    isConnect = checkConnect()
    if isConnect == False: exit()

    obj = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    kospiCodeList = obj.GetStockListByMarket(1)
    kosdaqCodeList = obj.GetStockListByMarket(0)

    codeList = kospiCodeList + kosdaqCodeList
    market = {1: 'KOSPI', 2: 'KOSDAQ', 3: 'K-OTC', 4: 'KRX', 5: 'KONEX'}

    stockList = dict()
    for i, code in enumerate(codeList):
        secondCode = obj.GetStockSectionKind(code)
        name = obj.CodeToName(code)
        stdPrice = obj.GetStockStdPrice(code)
        stockList[i] = (market[secondCode], name, stdPrice)

    return stockList