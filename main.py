import sys
import time
import ClassMoeum
from PyQt5.QtCore import QEventLoop
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *

class MyWindow(QMainWindow):
    cl = ClassMoeum

    def __init__(self):
        super().__init__()

        # Kiwoom Login
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        #self.kiwoom.dynamicCall("CommConnect()")
        self.kiwoom.CommConnect()

        # Open API+ Event

        # 클래스로 접근법
        # self.kiwoom.OnEventConnect.connect(self.cl.event_connect)
        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)
        self.kiwoom.OnReceiveTrCondition.connect(self.receive_trCondition)
        self.kiwoom.OnReceiveMsg.connect(self.receive_msg)
        self.kiwoom.OnReceiveChejanData.connect(self.receive_Chejan)

        self.setWindowTitle("PyStock")
        self.setGeometry(600, 600, 600, 600)

        label = QLabel('종목코드: ', self)
        label.move(20, 20)

        self.code_edit = QLineEdit(self)
        self.code_edit.move(80, 20)
        self.code_edit.setText("039490")

        self.text_edit = QTextEdit(self)
        self.text_edit.setGeometry(10, 140, 500, 180)
        self.text_edit.setEnabled(False)

        self.listWidget = QListWidget(self)
        self.listWidget.setGeometry(10, 330, 500, 200)

        btn1 = QPushButton("조회", self)
        btn1.move(190, 20)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("계좌조회", self)
        btn2.move(300, 20)
        btn2.clicked.connect(self.btn2_clicked)

        btn3 = QPushButton("예수금상세현황요청", self)
        btn3.setGeometry(10, 10, 150, 30)
        btn3.move(410, 20)
        btn3.clicked.connect(self.btn3_clicked)

        btn4 = QPushButton("종목코드 얻기", self)
        btn4.move(10, 60)
        btn4.clicked.connect(self.btn4_clicked)

        btn5 = QPushButton("주식거래원요청", self)
        btn5.move(120, 60)
        btn5.clicked.connect(self.btn5_clicked)

        btn6 = QPushButton("조건검색 리스트", self)
        btn6.move(230, 60)
        btn6.clicked.connect(self.btn6_clicked)

        btn7 = QPushButton("조건검색", self)
        btn7.move(340, 60)
        btn7.clicked.connect(self.btn7_clicked)

        btn8 = QPushButton("조건검색중지", self)
        btn8.move(450, 60)
        btn8.clicked.connect(self.btn8_clicked)

        btn9 = QPushButton("주문", self)
        btn9.move(10, 100)
        btn9.clicked.connect(self.btn9_clicked)

    def event_connect(self, err_code):
        if err_code == 0:
            self.text_edit.append("로그인 성공")

    def btn1_clicked(self):
        code = self.code_edit.text()
        self.text_edit.setText("종목코드 : " + code)
        # self.text_edit.append("종목코드 : " + code)

        # SetInputValue
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")

        # 전일 일별주가 요청
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", "20220208")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_pre_req", "opt10086", 0,
                                "0102")

        # 금일 일별주가 요청
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", "20220209")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_today_req", "opt10086", 0,
                                "0102")

    def btn2_clicked(self):
        account_num = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"])
        server_info = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")
        account_cnt = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "ACCOUNT_CNT")

        self.text_edit.setText("계좌번호: \n" + account_num.strip().replace(";","\n"))
        self.text_edit.append("계좌갯수: " + account_cnt)
        if server_info == 0:
            self.text_edit.append("모의투자")
        else:
            self.text_edit.append("실거래서버")

    def btn3_clicked(self):
        account_num = self.code_edit.text()
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "계좌번호", account_num)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")

        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opw00001", "opw00001", 0, "0101")

    def btn4_clicked(self):
        try:
            ret = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ["0"])

            kospi_code_list = ret.split(';')
            kospi_code_name_list = []
            count = 0

            for x in kospi_code_list:
                kospi_name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [x])
                kospi_code_name_list.append(x + " : " + kospi_name)
                print(self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [x]))
                count = count + 1
            self.listWidget.addItems(kospi_code_name_list)
            self.text_edit.append("코스피 총 갯수 : "+str(len(kospi_code_name_list)))

        except Exception as e:
            print(e)

    def btn5_clicked(self):
        try:
            code = self.code_edit.text()

            # SetInputValue
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

            # CommRqData
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10002_req", "opt10002", 0, "0101")


        except Exception as e:
            print(e)

    def btn6_clicked(self):
    # 조건식 리스트 불러오기
        try:
            #isLoad = self.kiwoom.dynamicCall("GetConditionLoad()")
            isLoad = self.kiwoom.GetConditionLoad()
            if isLoad == 0:
                self.text_edit.append("조건식 요청 실패")


            data = self.kiwoom.dynamicCall("GetConditionNameList()")
            print(data)
            if data == "":
                self.text_edit.append("getConditionNameList(): 사용자 조건식이 없습니다.")

            conditionList = data.split(';')
            del conditionList[-1]

            self.listWidget.clear()
            self.listWidget.addItems(conditionList)

            self.conditionLoop = QEventLoop()
            self.conditionLoop.exec_()


        except Exception as e:
            print(e)

    def btn7_clicked(self):
    # 조건식 검색
        try:
            screenNo = "0150"
            conditionName = self.code_edit.text().split("^")[1]
            conditionIndex = self.code_edit.text().split("^")[0]
            isRealTime = "1"
            isRequest = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)",
                                     screenNo, conditionName, conditionIndex, isRealTime)

            print(isRequest)
            if isRequest == 0:
                self.text_edit.setText("조건검색 실패")
            if isRequest == 1:
                self.text_edit.setText("조건검색 접속 성공")
            self.conditionLoop = QEventLoop()
            self.conditionLoop.exec_()

        except Exception as e:
            print(e)

    def btn8_clicked(self):
        try:
            screenNo = "0150"
            conditionName = self.code_edit.text().split("^")[1]
            conditionIndex = self.code_edit.text().split("^")[0]
            self.kiwoom.dynamicCall("SendConditionStop(QString, QString, int)",
                                     screenNo, conditionName, conditionIndex)

            self.conditionLoop = QEventLoop()
            self.conditionLoop.exec_()

        except Exception as e:
            print(e)

    def btn9_clicked(self):
    # 조건식 검색
        try:
            code = self.code_edit.text().split(",")[0]
            orderAmount = self.code_edit.text().split(",")[1]
            orderPrice = self.code_edit.text().split(",")[2]
            self.text_edit.append("코드 : " + code)
            self.text_edit.append("수량 : " + orderAmount)
            self.text_edit.append("금액 : " + orderPrice)
            # 사용자구분명, 화면번호, 계좌번호 10자리, 주문유형 1~7, 종목코드(6자리), 주문수량, 주문가격, 거래구분 00~81, 원주문번호
            '''
            isRequest = \
                self.kiwoom.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                                "SendOrder_req", "0211", "5061614811",1,code,
                                             int(orderAmount), int(orderPrice), "00")
            '''

            isRequest = self.kiwoom.SendOrder("SendOrder_req", "0211", "5061614811",1,code,
                                             int(orderAmount), int(orderPrice), "00", "")

            if isRequest == 0:
                self.text_edit.setText("매수주문 성공")
            else:
                self.text_edit.setText("매수주문 실패 : " + isRequest)
            self.conditionLoop = QEventLoop()
            self.conditionLoop.exec_()
        except Exception as e:
            print(e)

    def receive_Chejan(self, sGubun, nItemCnt,sFIdList):
        #체결구분. 접수와 체결시 '0'값, 국내주식 잔고변경은 '1'값, 파생잔고변경은 '4'
        print("receiveChejanData 호출")
        print(sGubun)
        print(nItemCnt)
        print(sFIdList)

    def receive_msg(self, screenNo, rqname, trcode, msg):
        try:
            self.text_edit.append(msg)
        except Exception as e:
            print(e)


    def receive_trCondition(self, screenNo, codes, conditionName, conditionIndex, inquiry):
    # 조건검색 결과 리스트 불러오기
        try:
            if codes == "":
                return
            codeList = codes.split(';')
            del codeList[-1]
            self.text_edit.setText(', '.join(codeList))

    # 일별주가요청 opt10086
            for x in codeList:
                code = x
                # SetInputValue
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", "20220209")
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")

                # CommRqData
                self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_req", "opt10086", 0,
                                        "0102")

        except Exception as e:
            print(e)

        finally:
            self.conditionLoop.exit()

    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        try:

            print(rqname)

            if rqname == "opt10001_req":
                name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "종목명")
                volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "거래량")
                price = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "현재가")
                BPS = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "BPS")

                self.text_edit.append("종목명: " + name.strip())
                self.text_edit.append("거래량: " + volume.strip())
                self.text_edit.append("현재가: " + price.strip())
                self.text_edit.append("BPS: " + BPS.strip())

            if rqname == "opt10002_req":
                name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "종목명")

                bound = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "등락율")

                self.text_edit.append("종목명: " + name.strip())
                self.text_edit.append("bound: " + bound.strip())

                if price < BPS:
                    self.text_edit.append("종목명: " + name.strip())
                    self.text_edit.append("현재가: " + price.strip())

            if rqname == "opw00001":
                diposit = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "예수금")

                self.text_edit.setText("예수금: " + diposit.strip())


            if rqname == "opt10086_pre_req":
                self.preStart = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "시가")
                self.preHigh = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "고가")
                self.preLow = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "저가")
                self.preEnd = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "종가")
                self.preDay = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                              rqname, 0, "전일비")

                print(self.preStart)
                print(self.preHigh)
                print(self.preLow)
                print(self.preEnd)
                print(self.preDay)

            if rqname == "opt10086_today_req":
                todayStart = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "시가")
                todayHigh = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "고가")
                todayLow = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "저가")
                todayEnd = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "종가")
                todayDay = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                              rqname, 0, "전일비")

                print(todayStart)
                print(todayHigh)
                print(todayLow)
                print(todayEnd)
                print(todayDay)

                print(int(todayHigh) - int(self.preHigh))
                print(int(todayLow) - int(self.preLow))
                print(int(todayEnd) - int(self.preEnd))
                print(int(todayDay) - int(self.preDay))



        except Exception as e:
            print(e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
