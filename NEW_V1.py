import sys
from PyQt5.QtWidgets import QApplication
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop
import pandas as pd
import datetime as date
from PyQt5.QtCore import QTimer
import time

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")  # 키움증권 OpenAPI+ 모듈 로드

        # 로그인 이벤트 연결
        self.OnEventConnect.connect(self._handle_login)

        # 실시간 데이터 처리 이벤트 연결
        self.OnReceiveRealData.connect(self._handle_real_data)
        print("NEW이벤트 핸들러가 연결되었습니다.")

        self.interest_stocks = {}
        self.purchased_today = set()  # 오늘 구매한 종목 집합을 초기화합니다.
        self.last_purchased_date = None  # 마지막 구매 날짜 초기화
        self.purchased_prices = {}

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
        if code not in self.interest_stocks:
            return

        print(f"실시간 데이터 수신: 종목코드={code}, 실시간 타입={real_type}, 데이터={real_data}")
        # 실제로 필요한 데이터가 들어오는지 확인하기 위한 추가적인 로깅
        current_time = date.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"Data received at {current_time}")
        # print(f"[{current_time}] 실시간 데이터 수신: 종목코드={code}, 실시간 타입={real_type}, 데이터={real_data}")
        if code not in self.interest_stocks:
            print(f"{code} is not in the interest stocks list!")
            return

        if real_type == "주식체결":
            # real_current_price = int(self.dynamicCall("GetCommRealData(QString, int)", code, 10))  # 현재가 정보
            current_price = abs(int(self.dynamicCall("GetCommRealData(QString, int)", code, 10)))
            threshold_price = self.interest_stocks.get(code)  # 사용자가 설정한 기준가
            print(f"현재가: {current_price}, 기준가: {threshold_price}")
            # self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", 0, "2000")

            if current_price <= threshold_price:
                # 조건 충족 시 로깅
                print(f"매수 조건 충족. 종목코드={code}, 현재가={current_price}, 기준가={threshold_price}")
                self.buy_stock(code, current_price)
                self.purchased_today.add(code)  # 해당 종목을 오늘 구매한 종목 집합에 추가


            # 매수 로직 실행
            #
            # if current_price <= threshold_price:
            #     # 구매 로직
            #     print(f"{code}의 현재가 {current_price}는 기준금액 {threshold_price}보다 낮습니다. 구매를 진행합니다.")



    def get_account_number(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)", "ACCNO")
        account_number_list = account_list.split(';')
        print("계좌번호 리스트:", account_number_list[:-1])  # 마지막 항목은 빈 문자열이므로 제외하고 출력
        return account_number_list[:-1]

    def detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        print(self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "예수금상세현황요청", "opw00001", sPrevNext,
                         self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        print(self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가잔고내역요청", "opw00018", sPrevNext,
                         self.screen_my_info)

        self.detail_account_info_event_loop2.exec_()


    def buy_stock(self, code, price):
        account_number = "8061329111"
        order_type = 1
        quantity = 1
        price = price
        order_no = ""

        result = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                  ["신규매수", order_no, account_number, order_type, code, quantity, price, "00", ""])
        if result == 0:
            print(f"주문 성공: {code}")

            # 주문이 성공하면 해당 주식 종목을 관심 종목 목록에서 제거
            if code in self.interest_stocks:
                del self.interest_stocks[code]

            # 관심 종목 목록을 업데이트하여 다시 구독
            time.sleep(5)
            # self.subscribe()

        else:
            print(f"주문 실패: {code}, 오류 코드 {result}")

        # self.last_purchased_date = date.today() #확인 주석처리후 실행해보기.
        # self.purchased_today.add(code)


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
            print(self.account_num)

        # 데이터프레임 생성
        df = pd.DataFrame(data, columns=['Code', 'Name', 'Price'])
        print(df)  # 콘솔에 출력하거나 필요한 곳에 데이터 사용

        def evaluate_buy(self, code, current_price, purchased_price):
            """주식 구매 평가 함수"""
            evaluation_rate = (current_price - purchased_price) / purchased_price * 100
            for row in BUY_STRATEGY:
                if evaluation_rate <= row["구입 시 평가율"]:
                    # 해당 금액으로 주식 구매
                    amount_to_buy = row["단계별금액"]
                    self.buy_stock(code, current_price, amount_to_buy)
                    break

        def evaluate_sell(self, code, current_price, purchased_price):
            """주식 매도 평가 함수"""
            evaluation_rate = (current_price - purchased_price) / purchased_price * 100
            sell_points = [(5, 0.1), (9, 0.2), (30, 0.3)]
            for rate, fraction in sell_points:
                if evaluation_rate >= rate:
                    # 주식의 일정 비율을 매도
                    self.sell_stock(code, fraction)
                    print(f"매도 조건 충족. 종목코드={code}, 현재 평가율={evaluation_rate:.2f}%, 매도 비율={fraction * 100}%")
                    break

        def sell_stock(self, code, fraction):
            # 주식 매도 로직
            # ... 기존 코드 ...

            # 계좌 내 해당 종목의 보유 수량을 가져옵니다.
            quantity = self.get_stock_quantity(code)
            sell_quantity = round(quantity * fraction)

            # 수량이 0보다 클 경우에만 매도 주문을 진행합니다.
            if sell_quantity > 0:
                # ... 매도 주문 코드 ...
                print(f"종목코드 {code}의 {sell_quantity}주를 매도 주문합니다.")
            else:
                print(f"종목코드 {code}의 매도할 수량이 없습니다.")

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

            rows = self.dynamicCall("GetRepeatCnt(QString, QString)", sTrCode, sRQName)
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목번호")
                code = code.strip()[1:]  # 종목 코드 앞에 'A'가 붙어있으므로 제거
                purchase_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                  "매입가")
                self.purchased_prices[code] = int(purchase_price.strip())

            print(
                "계좌평가잔고내역요청 싱글데이터 : %s - %s - %s" % (total_buy_money, total_profit_loss_money, total_profit_loss_rate))


BUY_STRATEGY = [
    {"단계": 1, "구입 시 평가율": -1.73, "단계별금액": 44796},
    {"단계": 2, "구입 시 평가율": -2.63, "단계별금액": 86345},]
# ... [나머지 전략]

if __name__ == "__main__":
    app = QApplication(sys.argv)
    kiwoom = Kiwoom()
    kiwoom.login()

    # 관심 종목 및 기준금액 불러오기
    kiwoom.get_account_number()
    kiwoom.interest_stocks = kiwoom.load_interest_stocks_from_excel("C:/Users/gky12/Desktop/interest_stock.xlsx")

    print("Loaded interest stocks:", kiwoom.interest_stocks)

    print(kiwoom.interest_stocks)

    kiwoom.subscribe()  # 종목 구독

    print()

    result = app.exec_()  # 이벤트 루프 시작


    kiwoom.detail_account_info()
    print(f"Event loop returned: {result}")