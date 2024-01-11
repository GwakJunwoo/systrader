import win32com.client

def checkConnect():
    obj = win32com.client.Dispatch("CpUtil.CpCybos")
    conn = obj.IsConnect
    if conn == 10: 
        print('PLUS가 정상적으로 연결되지 않음')
        return False
    else: return True