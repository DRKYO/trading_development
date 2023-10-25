from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from PyQt5.QtTest import *
#
# from config.errorCode import *
# from config.kiwoomType import *

# from config import errorCode

import pandas as pd
from datetime import datetime as dt


class Kiwoom(QAxWidget):

    day = dt.today().strftime("%Y%m%d")
    print(day)

    def __init__(self):
        super().__init__()
        print("Kiwoom() class start.")


        ####### event loop를 실행하기 위한 변수 모음
        self.login_event_loop = QEventLoop()  # 로그인 요청용 이벤트 루프
        self.detail_account_info_event_loop = QEventLoop()  # 예수금 요청용 이벤트 루프
        self.detail_account_info_event_loop2 = QEventLoop()  # 잔고 요청용 이벤트 루프
        #########################################

        ####### 계좌 관련된 변수
        self.account_stock_dict = {}
        self.account_dict = {}
        self.account_num = None  # 계좌번호 담아줄 변수
        self.deposit = 0  # 예수금
        self.use_money = 0  # 실제 투자에 사용할 금액
        self.use_money_percent = 0.5  # 예수금에서 실제 사용할 비율
        self.output_deposit = 0  # 출력가능 금액
        self.total_profit_loss_money = 0  # 총평가손익금액
        self.total_profit_loss_rate = 0.0  # 총수익률(%)
        ########################################

        ####### 요청 스크린 번호
        self.screen_my_info = "2000"  # 계좌 관련한 스크린 번호
        self.screen_calculation_stock = "4000"  # 계산용 스크린 번호
        self.screen_real_stock = "5000"  # 종목별 할당할 스크린 번호
        self.screen_meme_stock = "6000"  # 종목별 할당할 주문용 스크린 번호
        self.screen_start_stop_real = "1000"  # 장 시작/종료 실시간 스크린 번호
        ########################################

        ######### 초기 셋팅 함수들 바로 실행
        self.get_ocx_instance()  # OCX 방식을 파이썬에 사용할 수 있게 반환해 주는 함수 실행
        self.event_slots()  # 키움과 연결하기 위한 시그널 / 슬롯 모음
        self.real_event_slot()  # 실시간 이벤트 시그널 / 슬롯 연결
        self.signal_login_commConnect()  # 로그인 요청 함수 포함
        self.get_account_info()  # 계좌번호 가져오기

        self.detail_account_info()  # 예수금 요청 시그널 포함

        self.detail_account_mystock()  # 계좌평가잔고내역 가져오기
        #########################################

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 레지스트리에 저장된 API 모듈 불러오기

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot)  # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot)  # 트랜잭션 요청 관련 이벤트

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널
        self.login_event_loop.exec_()  # 이벤트 루프 실행

    # def login_slot(self, err_code):
    #     # print(errors(err_code)[1])
    #
    #     # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
    #     self.login_event_loop.exit()

    def login_slot(self):

        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")  # 계좌번호 반환
        account_num = account_list.split(';')[0]  # a;b;c  [a, b, c]

        self.account_num = account_num

        print("계좌번호 : %s" % account_num)

    def detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", sPrevNext,
                         self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext,
                         self.screen_my_info)

        self.detail_account_info_event_loop2.exec_()

    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):

        if sRQName == "예수금상세현황요청":
            self.account_dict[1] = {}
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "주문가능금액")
            self.deposit = int(deposit.strip())

            use_money = float(self.deposit) * self.use_money_percent
            self.use_money = int(use_money)
            self.use_money = self.use_money / 4

            output_deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0,
                                              "출금가능금액")
            self.output_deposit = int(output_deposit.strip())

            print("주문가능금액 : %s" % self.deposit)
            print("출금가능금액 : %s" % self.output_deposit)

            self.account_dict[1].update({"주문가능금액": self.deposit})
            self.account_dict[1].update({"출금가능금액": self.output_deposit})




            self.stop_screen_cancel(self.screen_my_info)

            self.detail_account_info_event_loop.exit()

        elif sRQName == "계좌평가잔고내역요청":
            total_buy_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0,
                                               "총매입금액")  # 출력 : 000000000746100
            self.total_buy_money = int(total_buy_money.strip())
            total_profit_loss_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,
                                                       0, "총평가손익금액")  # 출력 : 000000000009761
            self.total_profit_loss_money = int(total_profit_loss_money.strip())
            total_profit_loss_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,
                                                      0, "총수익률(%)")  # 출력 : 000000001.31
            self.total_profit_loss_rate = float(total_profit_loss_rate.strip())

            self.account_dict[1].update({"총매입금액": self.total_buy_money})
            self.account_dict[1].update({"총평가손익금액": self.total_profit_loss_money})
            self.account_dict[1].update({"총수익률(%)": self.total_profit_loss_rate})


            print(
                "계좌평가잔고내역요청 싱글데이터 : %s - %s - %s" % (total_buy_money, total_profit_loss_money, total_profit_loss_rate))

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]

                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,i, "매입금액")
                total_current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,i, "평가금액")
                today_buy_cnt = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,i, "금일매수수량")
                today_sell_cnt = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,i, "금일매도수량")
                yesterday_buy_cnt = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,i, "전일매수수량")
                yesterday_sell_cnt = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,i, "전일매도수량")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,"매매가능수량")


                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] = {}

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())
                total_current_price = int(total_current_price.strip())
                today_buy_cnt = int(today_buy_cnt.strip())
                today_sell_cnt = int(today_sell_cnt.strip())
                yesterday_buy_cnt = int(yesterday_buy_cnt.strip())
                yesterday_sell_cnt = int(yesterday_sell_cnt.strip())

                self.account_stock_dict[code].update({"종목코드": code})
                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({"평가금액": total_current_price})
                self.account_stock_dict[code].update({"금일매수수량": today_buy_cnt})
                self.account_stock_dict[code].update({"금일매도수량": today_sell_cnt})
                self.account_stock_dict[code].update({"전일매수수량": yesterday_buy_cnt})
                self.account_stock_dict[code].update({"전일매도수량": yesterday_sell_cnt})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})
                self.account_stock_dict[code].update({"날짜": self.day})

                print("종목번호: %s - 종목명: %s - 보유수량: %s - 매입가:%s - 수익률: %s - 현재가: %s" % (
                    code, code_nm, stock_quantity, buy_price, learn_rate, current_price))

            print("sPreNext : %s" % sPrevNext)
            print("계좌에 가지고 있는 종목은 %s " % rows)


            if sPrevNext == "2":
                self.detail_account_mystock(sPrevNext="2")
                self.a = pd.DataFrame.from_dict(self.account_stock_dict).T
                self.a.to_excel("C:/Users/USER/Desktop/price_crawling/g_new_update.xlsx")
                print("nice!!")
            else:

                self.a = pd.DataFrame.from_dict(self.account_stock_dict).T
                self.a.to_excel("C:/Users/USER/Desktop/price_crawling/g_new_update.xlsx")

                self.b = pd.DataFrame.from_dict(self.account_dict).T
                self.b.to_excel("C:/Users/USER/Desktop/price_crawling/stock_money.xlsx")

                self.detail_account_info_event_loop.exit()
            print("nice!!")

    def stop_screen_cancel(self, sScrNo=None):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)  # 스크린 번호 연결 끊기

    def chejan_slot(self, sGubun, nItemCnt, sFidList):

        if int(sGubun) == 0:  # 주문체결
            self.a.to_excel("C:/Users/USER/Desktop/price_crawling/stock_account_old.xlsx")
            self.stop_screen_cancel()

            self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
            self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
            self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
            self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")

            sPrevNext = "0"
            self.test = "2017"

            self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext,
                             self.test)
            self.detail_account_info_event_loop2.exec_()

    def real_event_slot(self):
        self.OnReceiveChejanData.connect(self.chejan_slot)  # 종목 주문체결 관련한 이벤트
