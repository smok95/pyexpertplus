"""야간선물옵션(EUREX) 조회 예제
"""
from urllib.request import Request
import pyexpertplus as ep

#  데이터 수신 이벤트 
class RequestDataEvent:
    def OnStatus(self, status, trCode, msgCode, msgName):
        print("OnStatus: status={}, trCode={}, msgCode={}, msgName={}".format(status, trCode, msgCode, msgName))

    def OnReceiveData(self, trCode, value, valueList, nextFlag, selectCount, msgCode, msgName):
        print("OnReceiveData: trCode={}, value={}, valueList={}, nextFlag={}, selectCount={}, msgCode={}, msgName={}"
        .format(trCode, value, valueList, nextFlag, selectCount, msgCode, msgName))

'''
        values = ep.YFValues()
        values.SetValueData(self.GetKorValueHeader(trCode), value)

        print("예수금:" + str(int(values.GetNameValue("예수금"))))
        print("출금가능금액:" + str(int(values.GetNameValue("출금가능금액"))))
        print("손익금액:" + str(int(values.GetNameValue("손익금액합계"))))
'''


def main():
    # ExpertPlus 초기화
    if ep.initialize()== False:
        return

    rq = ep.YFGRequest(RequestDataEvent)
    rq.GlobalInit()
    
    rq.RequestInit()
    rq.RequestData("GL0003")
    
    # 프로그램 종료 방지 루프
    ep.loop()

main()