import win32com.client
import sys
from checkConnect import *
from PyQt5.QtWidgets import *

def getStockPriceRealTime():

    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()

class CpEvent:
    instance = None

    def OnReceived(self):
        timess = CpEvent.instance.GetHeaderValue(18)  # 초
        exFlag = CpEvent.instance.GetHeaderValue(19)  # 예상체결 플래그
        cprice = CpEvent.instance.GetHeaderValue(13)  # 현재가
        diff = CpEvent.instance.GetHeaderValue(2)  # 대비
        cVol = CpEvent.instance.GetHeaderValue(17)  # 순간체결수량
        vol = CpEvent.instance.GetHeaderValue(9)  # 거래량

        if (exFlag == ord('1')):  # 동시호가 시간 (예상체결)
            print("실시간(예상체결)", timess, "*", cprice, "대비", diff, "체결량", cVol, "거래량", vol)
        elif (exFlag == ord('2')):  # 장중(체결)
            print("실시간(장중 체결)", timess, cprice, "대비", diff, "체결량", cVol, "거래량", vol)

class CpStockCur:
    def Subscribe(self, code):
        self.obj = win32com.client.Dispatch("DsCbo1.StockCur")
        win32com.client.WithEvents(self.obj, CpEvent)
        self.obj.SetInputValue(0, code)
        CpEvent.instance = self.obj
        self.obj.Subscribe()

    def Unsubscribe(self):
        self.obj.Unsubscribe()

class CpStockMst:
    def Request(self, code):
        isConnect = checkConnect()
        if isConnect == False: exit()

        obj = win32com.client.Dispatch("DsCbo1.StockMst")
        obj.SetInputValue(0, code)
        obj.BlockRequest()

        # 현재가 통신 및 통신 에러 처리
        rqStatus = obj.GetDibStatus()
        rqRet = obj.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False
        
        # 현재가 정보 조회
        code = obj.GetHeaderValue(0)  # 종목코드
        name = obj.GetHeaderValue(1)  # 종목명
        time = obj.GetHeaderValue(4)  # 시간
        cprice = obj.GetHeaderValue(11)  # 종가
        diff = obj.GetHeaderValue(12)  # 대비
        open = obj.GetHeaderValue(13)  # 시가
        high = obj.GetHeaderValue(14)  # 고가
        low = obj.GetHeaderValue(15)  # 저가
        offer = obj.GetHeaderValue(16)  # 매도호가
        bid = obj.GetHeaderValue(17)  # 매수호가
        vol = obj.GetHeaderValue(18)  # 거래량
        vol_value = obj.GetHeaderValue(19)  # 거래대금

        print("코드 이름 시간 현재가 대비 시가 고가 저가 매도호가 매수호가 거래량 거래대금")
        print(code, name, time, cprice, diff, open, high, low, offer, bid, vol, vol_value)
        return True     
    
class MyWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PLUS API TEST")
        self.setGeometry(300, 300, 1000, 1000)
        self.isRq = False
        self.objStockMst = CpStockMst()
        self.objStockCur = CpStockCur()

        btn1 = QPushButton("요청 시작", self)
        btn1.move(200, 20)
        btn1.resize(200,100)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("요청 종료", self)
        btn2.move(200, 300)
        btn2.resize(200, 100)
        btn2.clicked.connect(self.btn2_clicked)

        btn3 = QPushButton("종료", self)
        btn3.move(200, 600)
        btn3.resize(200, 100)
        btn3.clicked.connect(self.btn3_clicked)

    def StopSubscribe(self):
        if self.isRq:
            self.objStockCur.Unsubscribe()
        self.isRq = False

    def btn1_clicked(self):
        testCode = "A000660"
        if (self.objStockMst.Request(testCode) == False):
            exit()

        # 하이닉스 실시간 현재가 요청
        self.objStockCur.Subscribe(testCode)

        print("빼기빼기================-")
        print("실시간 현재가 요청 시작")
        self.isRq = True

    def btn2_clicked(self):
        self.StopSubscribe()

    def btn3_clicked(self):
        self.StopSubscribe()
        exit()