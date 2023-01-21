from pykiwoom.kiwoom import *
import time
import pandas as pd
from datetime import datetime

kiwoom = Kiwoom()
kiwoom.CommConnect(block=True)
print("블록킹 로그인 완료")

now = datetime.now()
now=now.strftime('%H:%M:%S')
now = int(now.replace(':', ''))
print(now)

while now < 151800:
    # TR 요청 (연속조회)
    df = kiwoom.block_request("opt10080",
                              종목코드="122630",
                              틱범위=1,
                              수정주가구분=1,
                              output="주식분봉차트조회",
                              next=0)
    # print(df.head())
    print(now)
    print(df)
    time.sleep(1)


# # 매수주문
# # 삼성전자, 10주, 시장가주문 매수
# kiwoom.SendOrder("시장가매수", "0101", stock_account, 1, "005930", 10, 0, "03", "")

# #매도주문
# # 삼성전자, 10주, 시장가주문 매도
# kiwoom.SendOrder("시장가매도", "0101", stock_account, 2, "005930", 10, 0, "03", "")

