from ast import While
from fileinput import nextfile
from logging import error, log
import string
import win32com.client as com
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
def initialize(waitingTime=5000) -> bool:
    """KB증권 ExpertPlus 초기화

    Args:
        waitingTime (int, optional): 마스터초기화 대기시간. Defaults to 5000.

    Raises:
        err: _description_

    Returns:
        bool: 성공시 True
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
    
    # 마스터 초기화를 위해 대기
    wait(waitingTime)


    if realObj:
        del realObj
        realObj = None
    
    return True

###############################################################################
# YFReal 이벤트
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
# 실시간데이터
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

################################################################################
# YFRequestData 이벤트
class YFRequestDataEvent:
    def OnStatus(self, status, trCode, msgCode, msgName):
        """송수신시 에러가 발생한 경우 발생되는 Event

            YFGRequest의 경우 RequestAliveInfo 등록시
            msgCode = 9002(해외선물 서버연결), 9003(야간옵션,선물 서버연결) 실시간으로 서버연결 확인 가능
        Args:
            status (string): 메시지상태코드(0 : 정상, 1:송신에러, 2:수신에러)
            trCode (string): 요청한 trCode
            msgCode (string): 송수신시 발생되는 에러메시지 코드
            msgName (string): 송수신시 발생되는 에러메시지 내용
        """       
        print("OnStatus: status={}, trCode={}, msgCode={}, msgName={}".format(status, trCode, msgCode, msgName))

    def OnReceiveData(self, trCode, value, valueList, nextFlag, selectCount, msgCode, msgName):
        """서버에서 데이터 정상 수신시 발생되는 Event

        Args:
            trCode (string): 요청한 TrCode
            value (string): 서버에서 수신된 단일 데이터
            valueList (string): 서버에서 수신된 리스트 데이터
            nextFlag (integer): 다음데이터가 있는 경우 1, 없으면 0
            selectCount (integer): 해당데이터 조회 개수 (기본:1, 다음Data 처리를 한 경우 증가)
            msgCode (string): 서버에서 수신된 메시지 코드
            msgName (string): 서버에서 수신된 메시지 내용
        """
        print("OnReceiveData: trCode={}, value={}, valueList={}, nextFlag={}, selectCount={}, msgCode={}, msgName={}"
        .format(trCode, value, valueList, nextFlag, selectCount, msgCode, msgName))

################################################################################
# 조회
class YFRequestData:
    def __init__(self, requestDataEventClass=None):
        if requestDataEventClass==None:
            self.comObj = com.Dispatch("YFExpertPlus.YFRequestData")
        else:
            self.comObj = com.DispatchWithEvents("YFExpertPlus.YFRequestData", requestDataEventClass)
    
    def ComInit(self):
        """COM DLL 초기화 함수(반드시 처음 화면 로딩시 호출 필요)
        """
        self.comObj.ComInit()
    
    def RequestInit(self):
        """서버로 데이터 요청시 전송하는 데이터 Format초기화하는 함수이며 데이터 요청시 반드시 호출
        """
        self.comObj.RequestInit()

    def SetData(self, name, value):
        """서버 전송시 필요한 데이터를 입력할 때 사용

        Args:
            name (any): 정의된 필드명 (제공문서참조)
            value (any): 정의된 필드에 대한 값
                Ex) SetData("Code", "003450") 
        """
        self.comObj.SetData(name, value)
    
    def SetListData(self, index, name, value):
        """길이가 정해져 있지 않고 리스트형식의 데이터 입력시 사용

        Args:
            index (integer): 0부터 시작
            name (any): 정의된 필드명 (문서제공)
            value (any): 정의된 필드에 대한 값             
        """
        self.comObj.SetListData(index, name, value)

    def RequestData(self, trCode, nextFlag) -> bool:
        """서버로 입력한 데이터 전송

        Args:
            trCode (any): 요청 TrCode
            nextFlag (integer): 기본값:0, 다음데이터가 존재하면 nextFlag: 1

        Returns:
            bool: 성공시 True
        """
        return self.comObj.RequestData(trCode, nextFlag)
    
    def GetData(self, name) -> any:
        """서버 전송시 입력된 데이터를 필드명으로 가져오기

        Args:
            name (any): 정의된 필드명(제공문서참고)

        Returns:
            any: 전송시 입력된 데이터
        """
        return self.comObj.GetData(name)
    
    def GetCommInfo(self) -> any:
        """통신 접속 정보

        Returns:
            any: 
        """
        return self.comObj.GetCommInfo()
    
    def GetAccountNo(self) -> string:
        """계좌번호 전체를 문자열로 가져오기

        Returns:
            string: 
        """
        return self.comObj.GetAccountNo()

    def AccountCount(self) -> int:
        """보유한 계좌 수

        Returns:
            int: 계좌 수
        """
        return self.comObj.AccountCount()
    
    def AccountItem(self, index) -> any:
        """인덱스로 계좌번호 가져오기

        Args:
            index (int): 0부터 시작

        Returns:
            any: 계좌번호
        """
        return self.comObj.AccountItem(index)
    
    def AllFormatExcel(self):
        """서버 송수신 Format을 Excel로 출력
        """
        self.comObj.AllFormatExcel()
    
    def GetKorValueHeader(self, trCode) -> any:
        """수신된 단일 데이터 Format을 한글 필드명으로 출력

        Args:
            trCode (any): 정의된 TrCode(제공문서참고)

        Returns:
            any: 
        """
        return self.comObj.GetKorValueHeader(trCode)
    
    def GetValueHeader(self, trCode) -> any:
        """수신된 단일 데이터 Format을 영문 필드명으로 출력

        Args:
            trCode (any): 정의된 TrCode (제공문서 참고)

        Returns:
            any: 
        """
        return self.comObj.GetValueHeader(trCode)
    
    def GetKorValueListHeader(self, trCode) -> any:
        """수신된 리스트 데이터 Format을 한글 필드명으로 출력

        Args:
            trCode (any): 정의된 TrCode (제공문서 참고)

        Returns:
            any: 
        """
        return self.comObj.GetKorValueListHeader(trCode)
    
    def GetValueListHeader(self, trCode) -> string:
        """수신된 리스트 데이터 Format을 영문 필드명으로 출력

        Args:
            trCode (string): 정의된 TrCode(제공문서 참고)

        Returns:
            string: 
        """
        return self.comObj.GetValueListHeader(trCode)
    
    def GetAllCodeName(self, code) -> string:
        """모든 종목리스트에서 코드를 찾아 종목명 출력

        Args:
            code (string): 종목코드

        Returns:
            string: 종목코드
        """
        return self.comObj.GetAllCodeName(code)
    
    def GetAllCodeType(self, code) -> string:
        """모든 종목리스트에서 코드를 찾아 종목 구분 출력
        데이터 조회시 종목 구분값을 입력해야 하는 경우 사용
        0: KOSPI, 1:KOSDAQ, 2:지수선물, D:스타지수선물, S:주식선물, 3:지수옵션, 4:ELW

        Args:
            code (string): 종목코드

        Returns:
            string: 종목구분
        """
        return self.comObj.GetAllCodeType(code)
    
    def GetCodeName(self, code, type) -> string:
        """코드에 대한 종목명을 리턴

        Args:
            code (string): 종목코드
            type (int): 시장구분
                1:주식, 2:선물, 3:옵션, 4:ELW, 5:스타지수선물, 6:주식선물

        Returns:
            string: 종목명
        """
        return self.comObj.GetCodeName(code, type)

    def GetCodeType(self, code, type) -> string:
        """코드에 대한 종목 구분 리턴

        Args:
            code (string): 종목코드
            type (int): 시장구분
                1:주식, 2:선물, 3:옵션, 4:ELW, 5:스타지수선물, 6:주식선물
        Returns:
            string: 종목구분
        """
        return self.comObj.GetCodeType(code, type)

    def GetElwStrCode(self, code) -> string:
        """ELW종목의 표준코드 리턴

        Args:
            code (string): 종목코드

        Returns:
            string: 표준코드
        """
        return self.comObj.GetElwStrCode(code)

    def GetAccountType(self, accountType) -> string:
        """계좌번호 전체를 문자열로 가져오기

        Args:
            accountType (int): 0-위탁계좌, 1-선물계좌, 9-미니원장계좌

        Returns:
            string: 
        """
        return self.comObj.GetAccountType(accountType)

    def GetMasterData(self):
        """서버로 마스터를 요청해서 다시 받아온다.
        """
        self.comObj.GetMasterData()
    
    def CheckMaster(self, type, count) -> int:
        """실제 종목수와 로컬에 받은 종목수를 TQ4008로 종목수를 조회 후 받은 값을 count에 입력

        Args:
            type (int): 0-코스피, 1-zhtmekr
            count (int): 실제 종목수

        Returns:
            int: 0-정상, 1-비정상
        """
        self.comObj.CheckMaster(type, count)
    
    def RequestAliveInfo(self, aType):
        """서버의 접속 상태를 확인한다.
            OnStatus 이벤트에 Status=3, msgCode=9004(연결), 9104(끊김)

        Args:
            aType (int): 11:통신정보수신
        """
        self.comObj.RequestAliveInfo(aType)

    def GetExpierMonth(self, aType):
        """옵션만기월 정보를 가져온다.

        Args:
            aType (int): 2-옵션 만기월 정보 코드
        """
        self.comObj.GetExpierMonth(aType)

    def RequestCommClose(self):
        """통신에 종료메시지를 보낸다.
        """
        self.comObj.RequestCommClose()
    
    def GSComInit(self, aType):
        """해외주식 서버에 접속을 하고 "실시간시세신청여부" 등록과 마스터 다운로드
        실시간시세 신청/해지는 "ACE[7451]해외주식 실시간시세 신청/해지"에서 가능합니다.

        Args:
            aType (int): 0-전체, 1-미국, 2-기타
        """
        self.comObj.GSComInit(aType)
    
    def GSRealReg(self, aType):
        """"실시간시세신청여부" 등록 후 마스터 다운로드
        해외주식 서버에 접속되어 있을 경우 사용 (기존에 리얼 등록한게 있으면 종료후 다시 등록해 주어야 합니다.)

        Args:
            aType (int): 0-전체, 1-미국, 2-기타
        """
        self.comObj.GSRealReg(aType)

################################################################################
# 단일데이터
class YFValues:
    def __init__(self):        
        self.comObj = com.Dispatch("YFExpertPlus.YFValues")
        self.Delimiter = self.comObj.Delimiter
        self.Data = self.comObj.Data

    def SetValueData(self, header, data):
        """수신된 단일 데이터를 해당 객체에 입력하는 함수
        YFRequestData 객체의 ReceiveData Event에서 해당 객체를 처리한다.
        ReceiveData Event에서 Value로 넘어온 데이터를 처리하는데 사용된다.  

        Args:
            header (Variant): 헤더정보
            data (Variant): 데이터            
        """
        self.comObj.SetValueData(header, data)

    def GetColCount(self):
        """해당 객체의 Col 개수

        Returns:
            integer: 해당 객체의 Col 개수
        """
        return self.comObj.GetColCount()
    
    def GetValue(self, index):
        """해당 컬럼의 값을 리턴

        Args:
            index (integer): 컬럼 index 0부터 시작
        """
        return self.comObj.GetValue(index)

    def SetValue(self, index, value):
        """해당 컬럼에 값을 변경

        Args:
            index (integer): 컬럼 index (0부터 시작)
            value (any): 변경할 값
        """
        self.comObj.SetValue(index, value)
    
    def GetNameValue(self, name):
        """해당 필드명에 위치한 값 리턴

        Args:
            name (string): 컬럼 명

        Returns:
            any : 해당 필드명에 위치한 값
        """
        return self.comObj.GetNameValue(name)
    
    def SetNameValue(self, name, value):
        """해당 컬럼명에 값 입력

        Args:
            name (string): 컬럼명
            value (any): 변경할 컬럼 값
        """
        self.comObj.SetNameValue(name, value)



################################################################################
# Global 조회
class YFGRequest:
    def __init__(self, requestDataEventClass=None):
        PROG_ID = "YFGExpertPlus.YFGRequest"
        if requestDataEventClass==None:
            self.comObj = com.Dispatch(PROG_ID)
        else:
            self.comObj = com.DispatchWithEvents(PROG_ID, requestDataEventClass)
    
    def GlobalInit(self):
        """COM DLL 초기화 함수(반드시 처음 화면 로딩시 호출 필요)
        """
        self.comObj.GlobalInit()
    
    def RequestInit(self):
        """서버로 데이터 요청시 전송하는 데이터 Format초기화하는 함수이며 데이터 요청시 반드시 호출
        """
        self.comObj.RequestInit()
    
    def SetData(self, name, value):
        """서버 전송시 필요한 데이터를 입력할 때 사용

        Args:
            name (any): 정의된 필드명 (제공문서참조)
            value (any): 정의된 필드에 대한 값

            Ex) SetData("Code", "6AM13")
        """
        self.comObj.SetData(name, value)

    def RequestData(self, trCode, nextFlag=0) -> bool:
        """서버로 입력한 데이터 전송

        Args:
            trCode (any): 요청 TrCode
            nextFlag (int): 디폴트:0 다음데이터가 존재하면 nextFlag:1

            Ex) 데이터 요청시 예
            RequestInit <- 전송초기화
            SetData("Code", "6AM13") <- 데이터 요청시 필요한 데이터
            RequestData("GQ9001", 0) <- 서버로 데이터 요청

        Returns:
            bool: 
        """
        return self.comObj.RequestData(trCode, nextFlag)
    
    def GetData(self, name) -> any:
        """서버 전송시 입력된 데이터를 필드명으로 가져오기

        Args:
            name (any): 정의된 필드명 (제공문서참조)

        Returns:
            any: 
        """
        return self.comObj.GetData(name)
    
    def GetCommInfo(self) -> any:
        """통신 접속 정보

        Returns:
            any: 
        """
        return self.comObj.GetcommInfo()

    def GetAccountNo(self) -> any:
        """계좌번호 전체를 문자열로 가져오기

        Returns:
            any: 
        """
        return self.comObj.GetAccountNo()
    
    def AccountCount(self) -> int:
        """보유한 계좌 개수

        Returns:
            int: 
        """
        return self.comObj.AccountCount()
    
    def AccountItem(self, index) -> any:
        """인덱스로 계좌번호 가져오기

        Args:
            index (int): 0부터 시작

        Returns:
            any: 
        """
        self.comObj.AccountItem(index)
    
    def AllFormatExcel(self):
        """서버 송수신 Format을 Excel로 출력
        """
        self.comObj.AllFormatExcel()
    
    def GetKorValueHeader(self, trCode) -> any:
        return self.comObj.GetKorValueHeader(trCode)
    
    def GetValueHeader(self, trCode) -> any:
        return self.comObj.GetValueHeader(trCode)
    
    def GetKorValueListHeader(self, trCode) -> any:
        return self.comObj.GetKorValueListHeader(trCode)
    
    def GetValueListHeader(self, trCode) -> string:
        return self.comObj.GetValueListHeader(trCode)
    
    def GetAllCodeName(self, code) -> any:
        return self.comObj.GetAllCodeName(code)
    
    def GetAccountType(self, accountType) -> any:
        return self.comObj.GetAccountType(accountType)
    
    def GetGFormatValue(self, aType, aValue) -> string:
        return self.comObj.GetGFormatValue(aType, aValue)
    
    def RequestAliveInfo(self, aType):
        self.comObj.RequestAliveInfo(aType)
    
    def GetHogaData(self, code, sValue) -> any:
        return self.comObj.GetHogaData(code, sValue)
    
    def RequestCommClose(self):
        self.comObj.RequestCommClose()
    
    def GetExCodeToExName(self, exCode)-> any:
        return self.comObj.GetExCodeToExName(exCode)
    

################################################################################
def unloop():
    global g_stopLoop
    g_stopLoop = True

def loop(callback=None, userdata=None):
    """메시지 펌프

    Args:
        callback (function, optional): loop이벤트 콜백. Defaults to None.
        userdata (any, optional): userdata. Defaults to None.
    """
    global g_stopLoop
    g_stopLoop = False
    while not g_stopLoop:
        pythoncom.PumpWaitingMessages()
        if callback == None:
            time.sleep(0.00001)
        else:
            callback(userdata)



def test():   
    real = YFReal(YFRealEvent)
    real.AddRealCode("000660", "RQ1101")
    real.AddRealCode("005930", "RQ1101")
    loop()


if __name__ == '__main__':    
    test()
    