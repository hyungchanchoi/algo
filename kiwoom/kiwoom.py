from PyQt5.QAxContainer import * 
from PyQt5.QtCore import *
from PyQt5.QtTest import *

from config.errorCode import *
import pandas as pd
import matplotlib.pyplot as plt 
from datetime import datetime

class Kiwoom(QAxWidget):
    def __init__(self) :
        super().__init__()
        
        print('kiwoom 클래스')
        
        ###eventloop
        self.login_event_loop = QEventLoop()  #None
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        ######################
        
        ###스크린 번호 모음
        self.screen_my_info = '2000'
        self.screen_calculation_stock = '4000'
        
        ###변수모음
        self.account_num = None
        ######################
        
        ###계좌 관련 변수
        self.use_money = 0
        self.use__money_percent = 1
        ######################
        
        ###변수모음
        self.account_stock_dict = {}

        self.get_ocx_instence()
        self.event_slots()
        
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()
        self.detail_account_mystock()
        self.not_concluded_account()
        self.day_kiwoom_db()
        
    def get_ocx_instence(self):
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')
        
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
        
    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))
        self.login_event_loop.exit()
        
    def signal_login_commConnect(self): 
        self.dynamicCall('CommConnect()')
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()
        
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        print(account_list)
        self.account_num = account_list.split(';')[0]
        print('나의 보유 계좌번호 %s' % self.account_num)
        
    def detail_account_info(self):
        print('예수금을 요청하는 부분')
        print(self.account_num)
        self.dynamicCall('SetInputValue(String, String)','계좌번호',self.account_num)
        self.dynamicCall('SetInputValue(String, String)','비밀번호','0917')
        self.dynamicCall('SetInputValue(String, String)','비밀번호입력매체구분','00')
        self.dynamicCall('SetInputValue(String, String)','조회구분','2')
        self.dynamicCall('CommRqData(String, String,int, String)','예수금상세현황요청','opw00001','0',self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
        
    def detail_account_mystock(self, sPrevNext = '0'):
        self.dynamicCall('SetInputValue(String, String)','계좌번호',self.account_num)
        self.dynamicCall('SetInputValue(String, String)','비밀번호','0917')
        self.dynamicCall('SetInputValue(String, String)','비밀번호입력매체구분','00')
        self.dynamicCall('SetInputValue(String, String)','조회구분','2')
        self.dynamicCall('CommRqData(String, String,int, String)','계좌평가잔고내역요청','opw00018',sPrevNext,self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
    
    def not_concluded_account(self, sPrevNext = '0'):
        
        self.dynamicCall('SetInputValue(QString, QString)','계좌번호',self.account_num)
        self.dynamicCall('SetInputValue(QString, QString)','체결구분','1')
        self.dynamicCall('SetInputValue(QString, QString)','매매구분','0')
        self.dynamicCall('CommRqData(String, String,int, String)','실시간미체결요청','opt10075',sPrevNext,self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
        
    
    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName,sPrevNext) :
        
        if sRQName == '예수금상세현황요청':
            deposit = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '예수금' ))
            print('예수금 : %s' % deposit)
            
            self.use__money = deposit * self.use__money_percent
            
            ok_deposit = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '출금가능금액' ))
            print('출금가능금액 : %s' % ok_deposit)
            
            self.detail_account_info_event_loop.exit()
            
        if sRQName == '계좌평가잔고내역요청':
            
            total_buy_money = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총매입금액' ))
            print('총매입금액 : %s' % total_buy_money)
            
            total_eval_money = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총평가금액' ))
            print('총평가금액 : %s' % total_eval_money)
            
            total_eval_profit = int(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총평가손익금액' ))
            print('총평가손익금액 : %s' % total_eval_profit)
    
            total_profit_rate = float(self.dynamicCall('GetCommData(String, String, int, String)', sTrCode, sRQName, 0, '총수익률(%)' ))
            print('총수익률 : %s' % total_profit_rate)
            
            rows = self.dynamicCall('GetRepeatCnt(QString, Qstring)',sTrCode,sRQName)
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall('GetCommData(QString, QString,int,Qstring)', sTrCode, sRQName, i,'종목번호')
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "보유수량")
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '종목명') 
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                earn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매매가능수량")
                
                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict.update({code:{}})
                    
                code = code.strip()[1:]
                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                earn_rate = float(earn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip)

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": earn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})
                
            if sPrevNext == '2':
                self.detail_account_mystock(sPrevNext='2')
            else:
                self.detail_account_info_event_loop.exit()

        if sRQName == '실시간미체결요청':
            
            rows = self.dynamicCall('GetRepeatCnt(QString, Qstring)',sTrCode,sRQName)
            
            for i in range(rows):
                code = self.dynamicCall('GetCommData(QString, QString,int,Qstring)', sTrCode, sRQName, i,'종목번호')
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, '종목명') 
                order_no = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문번호")
                order_status = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문상태") # 접수,확인,체결
                order_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문수량")
                order_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문가격")
                order_gubun = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "주문구분") # -매도, +매수, -매도정정, +매수정정
                not_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결량")
                            
                code = code.strip()
                code_nm = code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                order_quantity = int(order_quantity.strip())
                order_price = int(order_price.strip())
                order_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())
                
                if order_no in self.not_concluded_account:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}
                    
                self.not_account_stock_dict[order_no].update({'종목코드': code})
                self.not_account_stock_dict[order_no].update({'종목명': code_nm})
                self.not_account_stock_dict[order_no].update({'주문번호': order_no})
                self.not_account_stock_dict[order_no].update({'주문상태': order_status})
                self.not_account_stock_dict[order_no].update({'주문수량': order_quantity})
                self.not_account_stock_dict[order_no].update({'주문가격': order_price})
                self.not_account_stock_dict[order_no].update({'주문구분': order_gubun})
                self.not_account_stock_dict[order_no].update({'미체결수량': not_quantity})
                self.not_account_stock_dict[order_no].update({'체결량': ok_quantity})
                
            self.detail_account_info_event_loop.exit()
             
        if sRQName == '주식분봉차트조회요청':
            print('분몽데이터 요청')
            min_price = self.dynamicCall('GetCommData(QString, QString,int,Qstring)', sTrCode, sRQName, 0,'현재가')
            min_price = min_price.strip()
            min_time = self.dynamicCall('GetCommData(QString, QString,int,Qstring)', sTrCode, sRQName, 0,'체결시간')
            min_time = min_time.strip()

            cnt = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            
            min_price_list = []
            min_high_list = []
            min_low_list = []
            min_time_list = []
            
            now = datetime.now().date()
            print(now)
            
            for i in range(cnt):
                
                min_time = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "체결시간") 
                min_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "현재가")  
                min_high = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "고가") 
                min_low = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "저가") 
                
                min_time = min_time.strip()
                y = min_time[0:4]
                m = min_time[4:6]
                d = min_time[6:8]
                h = min_time[8:10]
                m_ = min_time[10:12]
                min_time = y + '-' + m + '-' + d + " " + h + ':' + m_ 
                min_time = datetime.strptime(min_time, '%Y-%m-%d %H:%M')

                min_price_list.append(int(min_price))
                min_high_list.append(int(min_high))
                min_low_list.append(int(min_low))
                min_time_list.append(min_time)

            #가끔 가다 음수가 나와서 절대값으로 바꿔주는 코드
            min_price_list = list(map(abs,min_price_list))
            min_high_list = list(map(abs,min_high_list))
            min_low_list = list(map(abs,min_low_list))
                                    
            data = {
                'time' : min_time_list,
                'price' :  min_price_list,
                'high' : min_high_list,
                'low' : min_low_list }
                        
            df = pd.DataFrame(data)
            # df.index = min_time_list
            df = df.sort_index(ascending=False)

            df['M'] = (df['price'] + df['high'] + df['low']) / 3  #https://dipsy-encyclopedia.tistory.com/62
            df['m'] = df['M'].rolling(window=100).mean()
            df = df[-500:]
            df['d'] = (df['M'] - df['m']).abs()
            df['d'] = df['d'].rolling(window=100).mean()
            df = df[-400:]
            df['cci'] = (df['M'] - df['m']) / (0.015 * df['d'])
            
            print(df)
            
            ###cci 그래프 그리는 코드
            # plt.figure(figsize=(25,10))
            # df['cci'].plot()
            # plt.show()
            # df.to_excel('df.xlsx')
                        
            ###반복조회 할 때 사용하느 코드, 아마 백테스트 할 때 사용할 듯
            # if sPrevNext == '2':
            #     self.day_kiwoom_db(code = '122630', sPrevNext=sPrevNext)
            # else:
            #     self.calculator_event_loop.exit()
            
                      
    def calculator_fnc(self):
        pass #임시        
            
    def day_kiwoom_db(self, code=None, date=None, sPrevNext='0'):
        
        # QTest.qWait(3600) 
    
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", '122630')
        self.dynamicCall("SetInputValue(QString, QString)", "틱범위", '1')
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식분봉차트조회요청", "opt10080", sPrevNext, self.screen_calculation_stock)

        self.calculator_event_loop.exec_()