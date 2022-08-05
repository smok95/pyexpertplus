import win32com.client as com
import timer
import threading
import pythoncom
import time
import win32api

class YFRequestDataEventHandler:
    def Status(self, Status, TrCode, MsgCode, MsgName):
        print('onStatus' + MsgName)
    
class YFRealEventHandler:
    def ReceiveData(self, TrCode, Value, MsgCode, MsgName):
        print("real")


def fnTimer(timerId, time):
    print("timer call")
    realData = com.Dispatch("YFExpertPlus.YFReal")
    #realData = com.DispatchWithEvents("YFExpertPlus.YFReal", YFRealEventHandler)

    ret = realData.AddRealCode("005930", "RQ1101")

    timer.kill_timer(timerId)

def wait(msec):
    now = win32api.GetTickCount()
    while True:
        pythoncom.PumpWaitingMessages()
        if(win32api.GetTickCount() - now) > msec:
            break


def requestAccount():
    # 계좌목록조회
    accountData = com.Dispatch("YFExpertPlus.YFRequestData")
    accountData = com.DispatchWithEvents("YFExpertPlus.YFRequestData", YFRequestDataEventHandler)
    accountData.ComInit()

    print(accountData.AccountCount())
    for i in range(accountData.AccountCount()):
        print(accountData.AccountItem(i))

def fnTrd(arg):
    pythoncom.CoInitialize()

    print("create YFReal")
    #realData = com.DispatchWithEvents("YFExpertPlus.YFReal", YFRealEventHandler)

    realData = com.Dispatch("YFExpertPlus.YFReal")

    now = win32api.GetTickCount()
    while True:
        pythoncom.PumpWaitingMessages()
        if (win32api.GetTickCount()-now) > 1000:
            break

    handler = com.WithEvents(realData, YFRealEventHandler)

    ret = realData.AddRealCode("005930", "RQ1101")
    print("AddREalCode result=" + str(ret))

    while(True):
        pythoncom.PumpWaitingMessages()
        

class PyRealData:
    def __init__(self):
        self.comObj = None
    
    def set_object(self, obj):
        self.comObj = obj
        com.WithEvents(self.comObj, PyRealData)
    
    def OnStatus(self, Status, TrCode, MsgCode, MsgName):
        print("onStatus")

    
    def OnReceiveData(self, TrCode, Value, MsgCode, MsgName):
        print("OnReceiveData" + TrCode + "," + Value + "," + MsgName)

    def ReceiveData(self, TrCode, Value, MsgCode, MsgName):
        print("real")

    def Request(self):
        self.comObj.AddRealCode("005930", "RQ1101")

def test():
    pythoncom.CoInitialize()

    requestAccount()

    real = PyRealData()
    real.set_object(com.Dispatch("YFExpertPlus.YFReal"))
    wait(3000)

    real.Request()

    while(True):
        pythoncom.PumpWaitingMessages()
        #time.sleep(0.1)
    
print("begin test")
test()
print("exit program")