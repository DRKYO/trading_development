import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop
import pandas as pd
import datetime as date
from PyQt5.QtCore import QTimer

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 키움증권 OpenAPI+ 모듈 로드

        # 로그인 이벤트 연결
        self.OnEventConnect.connect(self._handle_login)

        # 실시간 데이터 처리 이벤트 연결
        self.OnReceiveRealData.connect(self._handle_real_data)
        print("이벤트 핸들러가 연결되었습니다.")

        self.interest_stocks = {}
        self.purchased_today = set()  # 오늘 구매한 종목 집합을 초기화합니다.
        self.last_purchased_date = None  # 마지막 구매 날짜 초기화

    def login(self):
        self.dynamicCall("CommConnect()")
        self.login_event_loop = QEventLoop()
        self.login_event_loop.exec_()

        if self.last_purchased_date is not None and self.last_purchased_date < date.today():
            self.reset_purchased()  # 오늘 구매한 종목 기록을 초기화

    def _handle_login(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        else:
            print("로그인 실패: ", err_code)
        self.login_event_loop.exit()




    def _handle_real_data(self, code, real_type, real_data):
        print(f"실시간 데이터 수신: 종목코드={code}, 실시간 타입={real_type}, 데이터={real_data}")
        # 실제로 필요한 데이터가 들어오는지 확인하기 위한 추가적인 로깅
        current_time = date.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"Data received at {current_time}")
        print(f"[{current_time}] 실시간 데이터 수신: 종목코드={code}, 실시간 타입={real_type}, 데이터={real_data}")

        if real_type == "주식체결":
            current_price = int(self.dynamicCall("GetCommRealData(QString, int)", code, 10))  # 현재가 정보
            threshold_price = self.interest_stocks.get(code)  # 사용자가 설정한 기준가
            print(f"현재가: {current_price}, 기준가: {threshold_price}")

            if current_price <= threshold_price:
                # 조건 충족 시 로깅
                print(f"매수 조건 충족. 종목코드={code}, 현재가={current_price}, 기준가={threshold_price}")
                # 매수 로직 실행

            if current_price <= threshold_price:
                # 구매 로직
                print(f"{code}의 현재가 {current_price}는 기준금액 {threshold_price}보다 낮습니다. 구매를 진행합니다.")
                self.buy_stock(code, current_price)
                self.purchased_today.add(code)  # 해당 종목을 오늘 구매한 종목 집합에 추가



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

    def buy_stock(self, code, price):
        account_number = "5163828010"  # 본인의 계좌번호
        order_type = 1  # 신규 매수
        quantity = 10  # 매수할 주식 수, 이 부분은 사용자의 전략에 따라 달라집니다.
        price = price  # 매수 희망 가격
        order_no = ""  # 주문 번호, 신규 주문의 경우 공백

        # 키움증권 OpenAPI의 SendOrder 함수를 호출하여 주문을 전송합니다.
        result = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                  ["신규매수", order_no, account_number, order_type, code, quantity, price, "00", ""])
        if result == 0:
            print(f"주문 성공: {code}")
        else:
            print(f"주문 실패: {code}, 오류 코드 {result}")

        self.last_purchased_date = date.today()
        self.purchased_today.add(code)  # 해당 종목을 오늘 구매한 종목 집합에 추가

    def load_interest_stocks_from_excel(self, file_path):
        # 엑셀 파일 불러오기
        df = pd.read_excel(file_path)

        # 종목 코드가 문자열로 올바르게 처리되도록 수정
        df['종목코드'] = df['종목코드'].apply(lambda x: str(x).zfill(6))  # 종목 코드를 6자리 문자열로 처리

        # '종목코드'를 key로, '기준금액'을 value로 하는 사전 생성
        interest_stocks = df.set_index('종목코드')['기준금액'].to_dict()

        return interest_stocks

    def subscribe(self):
        codes = ";".join(map(str, self.interest_stocks.keys()))
        print(f"구독할 종목 코드: {codes}")  # 종목 코드가 올바르게 출력되는지 확인
        new = 2
        # 실시간 데이터 요청
        self.dynamicCall("SetRealReg(QString, QString, QString, QString)", "1000", codes, "10;12;15", "0")
        print(f"Subscribed real-time market data of {codes}")
        print(new)


        print("실시간 데이터 구독 요청이 전송되었습니다.")  # 이 로그가 출력되는지 확인

    def update_stock_prices(self):
        # 코스피와 코스닥의 코드 리스트를 가져옵니다.
        kospi_codes = self.get_code_list_by_market("0")  # 코스피
        kosdaq_codes = self.get_code_list_by_market("10")  # 코스닥

        all_codes = kospi_codes + kosdaq_codes

        data = []
        for code in all_codes:
            code_name = self.get_master_code_name(code)
            self.get_stock_price(code)
            current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", "opt10001", "StockPrice", 0,
                                             "현재가")
            data.append([code, code_name, current_price.strip()])

        # 데이터프레임 생성
        df = pd.DataFrame(data, columns=['Code', 'Name', 'Price'])
        print(df)  # 콘솔에 출력하거나 필요한 곳에 데이터 사용

if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.login()
    # 관심 종목 및 기준금액 불러오기
    kiwoom.interest_stocks = kiwoom.load_interest_stocks_from_excel("C:/Users/User/Desktop/interest_stock.xlsx")
    print(kiwoom.interest_stocks)

    kiwoom.subscribe()  # 종목 구독

    # 타이머 설정: update_stock_prices 메서드가 10초마다 호출되도록 설정
    # timer = QTimer()
    # timer.timeout.connect(kiwoom.update_stock_prices)
    # timer.start(10000)  # 10초마다 업데이트

    result = app.exec_()  # 이벤트 루프 시작
    print(f"Event loop returned: {result}")