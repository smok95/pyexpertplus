from ast import While
from logging import error, log
import string
import win32com.client as com
import timer
import threading
import pythoncom
import time
import win32api
import ctypes

g_stopLoop = False

###############################################################################
def wait(msec):
    now = win32api.GetTickCount()
    while True:
        pythoncom.PumpWaitingMessages()
        if(win32api.GetTickCount() - now) > msec:
            break

###############################################################################
def initialize() -> bool:
    """
    KB증권 ExpertPlus 초기화
    """
    # 관리자권한 여부 확인
    if not ctypes.windll.shell32.IsUserAnAdmin():
        error("관리자권한이 필요합니다.")
        return False

    # COM 초기화
    pythoncom.CoInitialize()
    
    # Expert Plus 설치상태 확인
    try:        
        realObj = com.Dispatch("YFExpertPlus.YFReal")
    except pythoncom.com_error as err:        
        if err.hresult==-2147221005:
            error("ExpertPlus 설치상태를 확인해주세요")
            return False
        else:
            raise err
    
    # 마스터 초기화를 위해 약 5초간 대기
    wait(3000)


    if realObj:
        del realObj
        realObj = None
    
    return True

###############################################################################
class YFRealEvent:    
    def OnStatus(self, status, trCode, msgCode, msgName):
        """송수신시 에러가 발생한 경우 발생되는 Event

        Args:
            status (string): 메시지상태코드(0 : 정상, 1:송신에러, 2:수신에러)
            trCode (string): 요청한 trCode
            msgCode (string): 송수신시 발생되는 에러메시지 코드
            msgName (string): 송수신시 발생되는 에러메시지 내용
        """       
        print("OnStatus: status={}, trCode={}, msgCode={}, msgName={}".format(status, trCode, msgCode, msgName))

    def OnReceiveData(self, trCode, Value, msgCode, msgName):
        """서버에서 데이터 정상 수신시 발생되는 Event

        Args:
            trCode (string): 요청한 trCode
            Value (string): 서버에서 수신된 단일 데이터
            msgCode (string): 서버에서 수신된 메시지 코드
            msgName (string): 서버에서 수신된 메시지 내용
        """        
        print("OnReceiveData: trCode={}, Value={}, msgCode={}, msgName={}".format(trCode, Value, msgCode, msgName))

###############################################################################
class YFReal:
    def __init__(self, realEventClass):
        self.comObj = com.DispatchWithEvents("YFExpertPlus.YFReal",realEventClass)

    def AddRealCode(self, Code, trCode) -> bool:
        """실시간 시세 데이터를 수신 받기 위한 코드 등록

        Args:
            Code (string): 종목코드 또는 업종코드
            trCode (string): 정의된 trCode (제공문서참조)

        Returns:
            bool: 
        """
        return self.comObj.AddRealCode(Code, trCode)    
    
    def AddAccount(self, account, trCode) -> bool:
        """실시간 체결/미체결 데이터를 수신 받기 위한 계좌 등록

        Args:
            account (string): 계좌번호
            trCode (string): 정의된 trCode (제공문서참조)

        Returns:
            bool: 
        """
        return self.comObj.AddAccount(account, trCode)
    
    def RemoveAccount(self, account, trCode) -> bool:
        """실시간 체결/미체결 데이터 수신 등록 해제

        Args:
            account (string): 계좌번호
            trCode (string): 정의된 trCode (제공문서참조)

        Returns:
            bool: 
        """
        return self.comObj.RemoveAccount(account, trCode)

    def RemoveRealCode(self, code, trCode) -> bool:
        """실시간 시세 데이터 수신 등록 해제

        Args:
            code (string): 종목코드 또는 업종코드
            trCode (string): 정의된 TrCode (제공문서참조)

        Returns:
            bool: 
        """
        return self.comObj.RemoveRealCode(code, trCode)

    def AllDeleteReal(self):
        """실시간 등록된 모든 TR을 해제
        """
        self.comObj.AllDeleteReal()
    
    def AllFormatExcel(self):
        """서버 송수신 Format을 Excel로 출력
        """
        self.comObj.AllFormatExcel()
    
    def GetKorValueHeader(self, trCode):
        """수신된 단일 데이터 Format을 한글 필드명으로 출력

        Args:
            trCode (string): 정의된 TrCode(제공문서참조)
        """
        return self.comObj.GetKorValueHeader(trCode)

    def GetValueHeader(self, trCode):
        return self.comObj.GetValueHeader(trCode)

    def GetAllCodeName(self, code):
        return self.comObj.GetAllCodeName(code)
    
    def GetAllCodeType(self, code):
        return self.comObj.GetAllCodeType(code)
    
    def GetCodeName(self, code, type) -> string:
        """코드에 대한 종목명을 출력

        Args:
            code (string): 종목코드
            type (integer): 1:주식, 2.선물, 3.옵션, 4.ELW, 5.스타지수선물, 6.주식선물

        Returns:
            string: 종목명
        """
        return self.comObj.GetCodeName(code, type)
    
    def GetCodeType(self, code, type):
        return self.comObj.GetCodeType(code, type)
    
    def GetElwStrCode(self, code):
        return self.comObj.GetElwStrCode(code)

def unloop():
    global g_stopLoop
    g_stopLoop = True

def loop():
    global g_stopLoop
    g_stopLoop = False
    while not g_stopLoop:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.00001)



def test():   
    real = YFReal(YFRealEvent)
    real.AddRealCode("000660", "RQ1101")
    loop()


if __name__ == '__main__':    
    test()
    