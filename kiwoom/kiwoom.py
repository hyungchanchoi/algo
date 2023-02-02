from PyQt5.QAxContainer import * 
from PyQt5.QtCore import *
from PyQt5.QtTest import *

from config.errorCode import *
from config.kiwoomType import *
import pandas as pd
import matplotlib.pyplot as plt 
from datetime import datetime

class Kiwoom(QAxWidget):
    def __init__(self) :
        super().__init__()
        
        print('kiwoom class')
        
        self.realType = RealType()
        
        ###eventloop
        self.login_event_loop = QEventLoop()  #None
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        ######################
        
        ###스크린 번호 모음
        self.screen_start_stop_real = '1000'
        self.screen_my_info = '2000'
        self.screen_calculation_stock = '4000'
        self.screen_real_stock = '5000'
        self.screen_meme_stock = '6000'
        ######################
        
        ###변수모음
        self.account_num = None
        self.account_stock_dict = {}
        self.not_account_stock_dict = {}
        ######################
        
        ###계좌 관련 변수
        self.use_money = 0
        self.use__money_percent = 1
        ######################
        
        ###종목 정보 가져오기
        self.portfolio_stock_dict = {'122630':{'종목명' : 'KODEX 레버리지'}, 
                                     '252670':{'종목명' : 'KODEX 200선물인버스2X'}}
        self.jango_dict = {}
        ######################

        ###함수 모음
        self.get_ocx_instence()
        self.event_slots()
        self.real_event_slots()
        
        self.signal_login_commConnect()
        self.get_account_info()
        self.detail_account_info()
        self.detail_account_mystock()
        self.not_concluded_account()
        
        self.day_kiwoom_db() 
        
        self.screen_number_setting()
        
        self.dynamicCall('SetRealReg(Qstring,Qstring,Qstring,Qstring)',self.screen_start_stop_real, '', self.realType.REALTYPE['장시작시간']['장운영구분'], '0' )
        
        for code in self.portfolio_stock_dict.keys():
            screen_num = self.portfolio_stock_dict[code]['스크린번호']
            fids = self.realType.REALTYPE['주식체결']['체결시간']
            self.dynamicCall('SetRealReg(Qstring,Qstring,Qstring,Qstring)',screen_num, code, fids, '1' )
            print('실시간 등록 코드: %s, 스크린번호: %s, fid번호: %s' % (code, screen_num,  fids))             
        
        
    def get_ocx_instence(self):
        self.setControl('KHOPENAPI.KHOpenAPICtrl.1')
        
    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)
        self.OnReceiveTrData.connect(self.trdata_slot)
        
    def login_slot(self, errCode):
        print(errCode)
        print(errors(errCode))
        self.login_event_loop.exit()
        
    def real_event_slots(self):
        self.OnReceiveRealData.connect(self.realdata_slot)
        self.OnReceiveChejanData.connect(self.chejan_slot)
        self.OnReceiveMsg.connect(self.msg_slot)
        
    def signal_login_commConnect(self): 
        self.dynamicCall('CommConnect()')
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()
        
    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(String)", "ACCNO")
        self.account_num = account_list.split(';')[0]
        print('***나의 보유 계좌번호 : %s *** ' % self.account_num)
        
    def detail_account_info(self):
        print('***예수금상세현황요청***')
        self.dynamicCall('SetInputValue(String, String)','계좌번호',self.account_num)
        self.dynamicCall('SetInputValue(String, String)','비밀번호','0917')
        self.dynamicCall('SetInputValue(String, String)','비밀번호입력매체구분','00')
        self.dynamicCall('SetInputValue(String, String)','조회구분','2')
        self.dynamicCall('CommRqData(String, String,int, String)','예수금상세현황요청','opw00001','0',self.screen_my_info)
        
        self.detail_account_info_event_loop.exec_()
        
    def detail_account_mystock(self, sPrevNext = '0'):
        print('***계좌평가잔고내역요청***')
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
        
    def day_kiwoom_db(self, code=None, date=None, sPrevNext='0'):
            
        # QTest.qWait(3600) 
    
        self.dynamicCall("SetInputValue(QString, QString)", "종목코드", '122630')
        self.dynamicCall("SetInputValue(QString, QString)", "틱범위", '1')
        self.dynamicCall("SetInputValue(QString, QString)", "수정주가구분", "1")

        if date != None:
            self.dynamicCall("SetInputValue(QString, QString)", "기준일자", date)

        self.dynamicCall("CommRqData(QString, QString, int, QString)", "주식분봉차트조회요청", "opt10080", sPrevNext, self.screen_calculation_stock)

        self.calculator_event_loop.exec_() 
        
    def screen_number_setting(self):
        screen_overwrite = []
        
        #계좌평가잔고내역에 있는 종목들
        for code in self.account_stock_dict.keys():
            if code not in screen_overwrite:
                screen_overwrite.append(code)
                
        #미체결에 있는 종목들
        for order_number in self.not_account_stock_dict.keys():
            code = self.not_account_stock_dict[order_number]['종목코드']
            
            if code not in screen_overwrite:
                screen_overwrite.append(code)
                
        #포트폴리오에 담겨 있는 종목들
        for code in self.portfolio_stock_dict.keys():
            screen_overwrite.append(code)
        
        cnt = 0
        for code in screen_overwrite:
            
            temp_screen = int(self.screen_real_stock)
            meme_screen = int(self.screen_meme_stock)
            
            if cnt % 50 == 0:
                temp_screen += 1
                self.screen_real_stock = str(temp_screen)
                
            if cnt % 50 == 0:
                meme_screen += 1
                self.screen_meme_stock = str(meme_screen)
                
            if code in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({'스크린번호': str(self.screen_real_stock),'주문용스크린번호': str(self.screen_meme_stock)})
                
            elif code not in self.portfolio_stock_dict.keys():
                self.portfolio_stock_dict[code].update({code :{'스크린번호': str(self.screen_real_stock),'주문용스크린번호': str(self.screen_meme_stock)}})
                
            cnt += 1 
        print(self.portfolio_stock_dict)
        
    def realdata_slot(self, sCode, sRealType, sRealData):
        
        if sRealType == '장시작시간':
            fid = self.realType.REALTYPE[sRealType]['장운영구분']
            value = self.dynamicCall('GetCommData(Qstring, int)', sCode, fid)
            
            if value == '0':
                print('장 시작 전')
            elif value == '3':
                print('장 시작')
            elif value == "2":
                print('장 종료, 동시호가로 넘어감')
            elif value == "4":
                print('3시30분 장 종료') 
        
        elif sRealType == '주식체결':
            a = self.dynamicCall('GetCommRealData(Qstring, int)', sCode, self.realType.REALTYPE[sRealType]['체결시간'])
            b = self.dynamicCall('GetCommRealData(Qstring, int)', sCode, self.realType.REALTYPE[sRealType]['현재가'])       
            b = abs(int(b)) 
            
            if sCode not in self.portfolio_stock_dict:
                self.portfolio_stock_dict.update({sCode:{}})
                 
            self.portfolio_stock_dict[sCode].update({'체결시간': a})
            self.portfolio_stock_dict[sCode].update({'현재가': b})
            
            #계좌잔고평가내역에 있고 오늘 산 잔고에는 없을 경우
            if sCode in self.account_stock_dict.keys() and sCode not in self.jango_dict.keys():
                
                print('신규매도 ㄱㄱ')
                
                meme_rate = (b - self.portfolio_stock_dict['매입가']) / self.portfolio_stock_dict['매입가'] * 100
                
                if self.portfolio_stock_dict['매매가능수량'] > 0 and (meme_rate > 5 or meme_rate < -1):
                    
                    order_success = self.dynamicCall('SendOrder(Qstring, Qstring, Qstring, int, Qstring, int, int, Qstring, Qstring)',
                                 ['신규매도',self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num,2,
                                 sCode, self.portfolio_stock_dict['매매가능수량'],0,self.realType.SENDTYPE['거래구분']['시장가'],''])
                    
                    if order_success == 0:
                        print('매도주문 전달 성공')
                        del self.account_stock_dict[sCode] 
                    else :
                        print('매도주문 전달 실패')
                
                
            elif sCode in self.jango_dict.keys():
                
                print('신규매도 ㄱㄱ')
            
                meme_rate = (b - self.portfolio_stock_dict['매입단가']) / self.portfolio_stock_dict['매매입단가입가'] * 100
                
                if self.portfolio_stock_dict['주문가능수량'] > 0 and (meme_rate > 5 or meme_rate < -1):
                    
                    order_success = self.dynamicCall('SendOrder(Qstring, Qstring, Qstring, int, Qstring, int, int, Qstring, Qstring)',
                                 ['신규매도',self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num,2,
                                 sCode, self.portfolio_stock_dict['매매가능수량'],0,self.realType.SENDTYPE['거래구분']['시장가'],''])
                    
                    if order_success == 0:
                        print('매도주문 전달 성공')
                        del self.account_stock_dict[sCode] 
                    else :
                        print('매도주문 전달 실패')
            
            elif sCode not in self.jango_dict:
                
                print('신규매수 ㄱㄱ')
                
                # result = (self.use_money * 0.1) / e 
                # quantuty = int(result)
                                
                order_success = self.dynamicCall('SendOrder(Qstring, Qstring, Qstring, int, Qstring, int, int, Qstring, Qstring)',
                                ['신규매수',self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num,1,
                                sCode, self.portfolio_stock_dict['매매가능수량'],0,self.realType.SENDTYPE['거래구분']['지정가'],''])
                
                if order_success == 0:
                    print('매수주문 전달 성공')
                    del self.account_stock_dict[sCode] 
                else :
                    print('매수주문 전달 실패')
                        
            not_meme_list = list(self.not_account_stock_dict)
            
            for order_num in not_meme_list:
                code = self.not_account_stock_dict[order_num]['종목코드']
                meme_price = self.not_account_stock_dict[order_num]['주문가격']
                not_quantity = self.not_account_stock_dict[order_num]['미체결수량']
                order_gubun = self.not_account_stock_dict[order_num]['주문구분']
                
                if order_gubun == '매수' and not_quantity > 0 and 0 > meme_price:
                    order_success = self.dynamicCall('SendOrder(Qstring, Qstring, Qstring, int, Qstring, int, int, Qstring, Qstring)',
                                ['매수취소', self.portfolio_stock_dict[sCode]['주문용스크린번호'], self.account_num, 3,
                                sCode, 0 , 0, self.realType.SENDTYPE['거래구분']['지정가'],'order_num'])
 
                    if order_success == 0:
                            print('매수취소 전달 성공')
                    else :
                        print('매수취소 전달 실패')
                    
                elif not_quantity == 0:
                    del self.not_account_stock_dict[order_num]
                    
    def chejan_slot(self, sGubun, nItemCnt, sFidList):
         
        if int(sGubun) == 0: #주문체결
            account_num = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['종목코드'])[1:]
            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['종목명'])
            stock_name = stock_name.strip()

            origin_order_number = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['원주문번호']) # 출력 : defaluse : "000000"
            order_number = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문번호']) # 출럭: 0115061 마지막 주문번호

            order_status = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문상태']) # 출력: 접수, 확인, 체결
            order_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문수량']) # 출력 : 3
            order_quan = int(order_quan)

            order_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문가격']) # 출력: 21000
            order_price = int(order_price)

            not_chegual_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['미체결수량']) # 출력: 15, default: 0
            not_chegual_quan = int(not_chegual_quan)

            order_gubun = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문구분']) # 출력: -매도, +매수
            order_gubun = order_gubun.strip().lstrip('+').lstrip('-')

            chegual_time_str = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['주문/체결시간']) # 출력: '151028'

            chegual_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['체결가']) # 출력: 2110 default : ''
            if chegual_price == '':
                chegual_price = 0
            else:
                chegual_price = int(chegual_price)

            chegual_quantity = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['체결량']) # 출력: 5 default : ''
            if chegual_quantity == '':
                chegual_quantity = 0
            else:
                chegual_quantity = int(chegual_quantity)

            current_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['현재가']) # 출력: -6000
            current_price = abs(int(current_price))

            first_sell_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['(최우선)매도호가']) # 출력: -6010
            first_sell_price = abs(int(first_sell_price))

            first_buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['주문체결']['(최우선)매수호가']) # 출력: -6000
            first_buy_price = abs(int(first_buy_price))

            ######## 새로 들어온 주문이면 주문번호 할당
            if order_number not in self.not_account_stock_dict.keys():
                self.not_account_stock_dict.update({order_number: {}})

            self.not_account_stock_dict[order_number].update({"종목코드": sCode})
            self.not_account_stock_dict[order_number].update({"주문번호": order_number})
            self.not_account_stock_dict[order_number].update({"종목명": stock_name})
            self.not_account_stock_dict[order_number].update({"주문상태": order_status})
            self.not_account_stock_dict[order_number].update({"주문수량": order_quan})
            self.not_account_stock_dict[order_number].update({"주문가격": order_price})
            self.not_account_stock_dict[order_number].update({"미체결수량": not_chegual_quan})
            self.not_account_stock_dict[order_number].update({"원주문번호": origin_order_number})
            self.not_account_stock_dict[order_number].update({"주문구분": order_gubun})
            self.not_account_stock_dict[order_number].update({"주문/체결시간": chegual_time_str})
            self.not_account_stock_dict[order_number].update({"체결가": chegual_price})
            self.not_account_stock_dict[order_number].update({"체결량": chegual_quantity})
            self.not_account_stock_dict[order_number].update({"현재가": current_price})
            self.not_account_stock_dict[order_number].update({"(최우선)매도호가": first_sell_price})
            self.not_account_stock_dict[order_number].update({"(최우선)매수호가": first_buy_price})

        elif int(sGubun) == 1: #잔고
            account_num = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['계좌번호'])
            sCode = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['종목코드'])[1:]

            stock_name = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['종목명'])
            stock_name = stock_name.strip()

            current_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['현재가'])
            current_price = abs(int(current_price))

            stock_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['보유수량'])
            stock_quan = int(stock_quan)

            like_quan = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['주문가능수량'])
            like_quan = int(like_quan)

            buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['매입단가'])
            buy_price = abs(int(buy_price))

            total_buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['총매입가']) # 계좌에 있는 종목의 총매입가
            total_buy_price = int(total_buy_price)

            meme_gubun = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['매도매수구분'])
            meme_gubun = self.realType.REALTYPE['매도수구분'][meme_gubun]

            first_sell_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['(최우선)매도호가'])
            first_sell_price = abs(int(first_sell_price))

            first_buy_price = self.dynamicCall("GetChejanData(int)", self.realType.REALTYPE['잔고']['(최우선)매수호가'])
            first_buy_price = abs(int(first_buy_price))

            if sCode not in self.jango_dict.keys():
                self.jango_dict.update({sCode:{}})

            self.jango_dict[sCode].update({"현재가": current_price})
            self.jango_dict[sCode].update({"종목코드": sCode})
            self.jango_dict[sCode].update({"종목명": stock_name})
            self.jango_dict[sCode].update({"보유수량": stock_quan})
            self.jango_dict[sCode].update({"주문가능수량": like_quan})
            self.jango_dict[sCode].update({"매입단가": buy_price})
            self.jango_dict[sCode].update({"총매입가": total_buy_price})
            self.jango_dict[sCode].update({"매도매수구분": meme_gubun})
            self.jango_dict[sCode].update({"(최우선)매도호가": first_sell_price})
            self.jango_dict[sCode].update({"(최우선)매수호가": first_buy_price})
            
            if stock_quan == 0:
                del self.jango_dict[sCode]
                self.dynamicCall("SetRealRemove(Qstring, Qstring)", self.portfolio_stock_dict[sCode['스크린번호'], sCode])
            

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
                code = self.dynamicCall('GetCommData(QString, QString,int,Qstring)', sTrCode, sRQName, i,'종목코드')
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
            print('***분봉데이터 요청***')
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
            self.calculator_event_loop.exit()
                   
    #송수신 메세지 get
    def msg_slot(self, sScrNo, sRQName, sTrCode, msg):
        print("스크린: %s, 요청이름: %s, tr코드: %s --- %s" %(sScrNo, sRQName, sTrCode, msg))