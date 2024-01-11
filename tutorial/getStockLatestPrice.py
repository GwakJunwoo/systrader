import win32com.client
from checkConnect import *

def getStockLatestPrice(code):
    isConnect = checkConnect()
    if isConnect == False: exit()

    obj = win32com.client.Dispatch("DsCbo1.StockMst")
    obj.SetInputValue(0, code)
    obj.BlockRequest()

    # 통신상태 확인 및 에러처리
    rqStatus = obj.GetDibStatus()
    rqRet = obj.GetDibMsg1()
    print("통신상태", rqStatus, rqRet)
    if rqStatus != 0: exit()

    # 현재가 정보 조회
    scode = obj.GetHeaderValue(0)
    name = obj.GetHeaderValue(1)
    time = obj.GetHeaderValue(4)
    cprice = obj.GetHeaderValue(11)
    diff = obj.GetHeaderValue(12)
    open = obj.GetHeaderValue(13)
    high = obj.GetHeaderValue(14)
    low = obj.GetHeaderValue(15)
    offer = obj.GetHeaderValue(16)
    bid = obj.GetHeaderValue(17)
    vol = obj.GetHeaderValue(18)
    vol_value = obj.GetHeaderValue(19)

    exFlag = obj.GetHeaderValue(58)
    exPrice = obj.GetHeaderValue(55)
    exDiff = obj.GetHeaderValue(56)
    exVol = obj.GetHeaderValue(57)

    return (scode, name, time, cprice, diff, open, high, low, offer, bid, vol, vol_value, exFlag, exPrice, exDiff, exVol)