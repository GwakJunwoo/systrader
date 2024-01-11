from Observer import *
import win32com

class Data(ABC):
    def getData(self, data):
        return data
    
class RequestData(Data):
    def __init__(self):
        self.obj = win32com.client.Dispatch("CpSysDib.StockChart")
 
    # 차트 요청 - 기간 기준으로
    def RequestFromTo(self, code, fromDate, toDate, caller):
 
        self.obj.SetInputValue(0, code)  # 종목코드
        self.obj.SetInputValue(1, ord('1'))  # 기간으로 받기
        self.obj.SetInputValue(2, toDate)  # To 날짜
        self.obj.SetInputValue(3, fromDate)  # From 날짜
        #self.obj.SetInputValue(4, 500)  # 최근 500일치
        self.obj.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 날짜,시가,고가,저가,종가,거래량
        self.obj.SetInputValue(6, ord('D'))  # '차트 주기 - 일간 차트 요청
        self.obj.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.obj.BlockRequest()
 
        len = self.obj.GetHeaderValue(3)
 
        caller.dates = []
        caller.opens = []
        caller.highs = []
        caller.lows = []
        caller.closes = []
        caller.vols = []
        for i in range(len):
            caller.dates.append(self.obj.GetDataValue(0,i))
            caller.opens.append(self.obj.GetDataValue(1, i))
            caller.highs.append(self.obj.GetDataValue(2, i))
            caller.lows.append(self.obj.GetDataValue(3, i))
            caller.closes.append(self.obj.GetDataValue(4, i))
            caller.vols.append(self.obj.GetDataValue(5, i))
 
        return [caller.dates, caller.opens, caller.highs, caller.lows, caller.closes, caller.vols]
 
    # 차트 요청 - 최근일 부터 개수 기준
    def RequestDWM(self, code, dwm, count, caller):
 
        self.obj.SetInputValue(0, code)  # 종목코드
        self.obj.SetInputValue(1, ord('2'))  # 개수로 받기
        self.obj.SetInputValue(4, count)  # 최근 500일치
        self.obj.SetInputValue(5, [0, 2, 3, 4, 5, 8])  # 요청항목 - 날짜,시가,고가,저가,종가,거래량
        self.obj.SetInputValue(6, dwm)  # '차트 주기 - 일/주/월
        self.obj.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.obj.BlockRequest()
 
        len = self.obj.GetHeaderValue(3)
 
        caller.dates = []
        caller.opens = []
        caller.highs = []
        caller.lows = []
        caller.closes = []
        caller.vols = []
        caller.times = []
        for i in range(len):
            caller.dates.append(self.obj.GetDataValue(0, i))
            caller.opens.append(self.obj.GetDataValue(1, i))
            caller.highs.append(self.obj.GetDataValue(2, i))
            caller.lows.append(self.obj.GetDataValue(3, i))
            caller.closes.append(self.obj.GetDataValue(4, i))
            caller.vols.append(self.obj.GetDataValue(5, i))
 
        print(len)
 
        return
 
    # 차트 요청 - 분간, 틱 차트
    def RequestMT(self, code, dwm, count, caller):
        # 연결 여부 체크
        bConnect = g_objCpStatus.IsConnect
        if (bConnect == 0):
            print("PLUS가 정상적으로 연결되지 않음. ")
            return False
 
        self.obj.SetInputValue(0, code)  # 종목코드
        self.obj.SetInputValue(1, ord('2'))  # 개수로 받기
        self.obj.SetInputValue(4, count)  # 조회 개수
        self.obj.SetInputValue(5, [0, 1, 2, 3, 4, 5, 8])  # 요청항목 - 날짜, 시간,시가,고가,저가,종가,거래량
        self.obj.SetInputValue(6, dwm)  # '차트 주기 - 분/틱
        self.obj.SetInputValue(7, 1)  # 분틱차트 주기
        self.obj.SetInputValue(9, ord('1'))  # 수정주가 사용
        self.obj.BlockRequest()
 
        rqStatus = self.obj.GetDibStatus()
        rqRet = self.obj.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            exit()
 
        len = self.obj.GetHeaderValue(3)
 
        caller.dates = []
        caller.opens = []
        caller.highs = []
        caller.lows = []
        caller.closes = []
        caller.vols = []
        caller.times = []
        for i in range(len):
            caller.dates.append(self.obj.GetDataValue(0, i))
            caller.times.append(self.obj.GetDataValue(1, i))
            caller.opens.append(self.obj.GetDataValue(2, i))
            caller.highs.append(self.obj.GetDataValue(3, i))
            caller.lows.append(self.obj.GetDataValue(4, i))
            caller.closes.append(self.obj.GetDataValue(5, i))
            caller.vols.append(self.obj.GetDataValue(6, i))
 
        print(len)
 
        return