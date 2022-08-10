pyexpertplus 사용을 위한 기본 필수사항
- 파이썬3 32비트 버전
- pywin32

```python
import pyexpertplus as ep

# 실시간 데이터 수신 이벤트 
class RealEvent:    
    def OnReceiveData(self, trCode, value, msgCode, msgName):        
        print("OnReceiveData: trCode={}, value={}, msgCode={}, msgName={}".format(trCode, value, msgCode, msgName))


def main():
    # ExpertPlus 초기화
    if ep.initialize()== False:
        return

    # 실시간 조회 객체 생성
    real = ep.YFReal(RealEvent)

    # 실시간 조회
    real.AddRealCode("000660", "RQ1101")

    # 프로그램 종료 방지 루프
    ep.loop()


main()
```
