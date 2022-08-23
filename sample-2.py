"""계좌 잔고내역 조회 예제
"""
import pyexpertplus as ep

#  데이터 수신 이벤트 
class RequestDataEvent:
    def OnStatus(self, status, trCode, msgCode, msgName):
        print("OnStatus: status={}, trCode={}, msgCode={}, msgName={}".format(status, trCode, msgCode, msgName))

    def OnReceiveData(self, trCode, value, valueList, nextFlag, selectCount, msgCode, msgName):
        #print("OnReceiveData: trCode={}, value={}, valueList={}, nextFlag={}, selectCount={}, msgCode={}, msgName={}"
        #.format(trCode, value, valueList, nextFlag, selectCount, msgCode, msgName))

        values = ep.YFValues()
        values.SetValueData(self.GetKorValueHeader(trCode), value)

        print("예수금:" + str(int(values.GetNameValue("예수금"))))
        print("출금가능금액:" + str(int(values.GetNameValue("출금가능금액"))))
        print("손익금액:" + str(int(values.GetNameValue("손익금액합계"))))


def main():
    # ExpertPlus 초기화
    if ep.initialize()== False:
        return

    accountData = ep.YFRequestData(RequestDataEvent)
    accountData.ComInit()

    accountData.RequestInit()
    accNo = "00000000000"   # 계좌번호
    accPwd = "0000"         # 계좌비밀번호
    accountData.SetData("Account", accNo)
    accountData.SetData("Password", accPwd)
    accountData.RequestData("TA1001", 0)
    
    # 프로그램 종료 방지 루프
    ep.loop()

main()