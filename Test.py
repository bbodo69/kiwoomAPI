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

        # 이벤트루프 선언
        self.event_connect_loop = QEventLoop()
        self.receive_trCondition_loop = QEventLoop()

        # Open API+ Event

        # 클래스로 접근법
        # self.kiwoom.OnEventConnect.connect(self.cl.event_connect)
        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.event_connect_loop.exec_()
        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)
        self.kiwoom.OnReceiveTrCondition.connect(self.receive_trCondition)
        self.kiwoom.OnReceiveMsg.connect(self.receive_msg)
        self.kiwoom.OnReceiveChejanData.connect(self.receive_Chejan)

        lstCode = ['900310', '900340', '005110', '900310', '900340', '005110', '900310', '900340', '005110']

        for code in lstCode:
            # 전일 일별주가 요청
            print(code)
        # 전일 일별주가 요청
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", "20220214")
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")

            # CommRqData
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_pre_req", "opt10086", 0,
                                    "0102")
            self.receive_trCondition_loop.exec()


    def event_connect(self, err_code):
        if err_code == 0:
            print("로그인 성공")
        self.event_connect_loop.exit()


    def btn1_clicked(self):
        code = self.code_edit1.text()
        self.text_edit.setText("종목코드 : " + code)
        # self.text_edit.append("종목코드 : " + code)

        # SetInputValue
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")

        lstCode = ['900310', '900340']

        for code in lstCode:
            # 전일 일별주가 요청
            print(code)
            preDay = self.code_edit2.text()
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", preDay)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")

            # CommRqData
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_pre_req", "opt10086", 0,
                                    "0102")


            # 금일 일별주가 요청
            todayDay = self.code_edit3.text()
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", todayDay)
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
        account_num = self.code_edit1.text()
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
            code = self.code_edit1.text()

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
            conditionName = self.code_edit1.text().split("^")[1]
            conditionIndex = self.code_edit1.text().split("^")[0]
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
            conditionName = self.code_edit1.text().split("^")[1]
            conditionIndex = self.code_edit1.text().split("^")[0]
            self.kiwoom.dynamicCall("SendConditionStop(QString, QString, int)",
                                     screenNo, conditionName, conditionIndex)

            self.conditionLoop = QEventLoop()
            self.conditionLoop.exec_()

        except Exception as e:
            print(e)

    def btn9_clicked(self):
    # 주문
        try:
            code = self.code_edit1.text().split(",")[0]
            orderAmount = self.code_edit1.text().split(",")[1]
            orderPrice = self.code_edit1.text().split(",")[2]
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
            print(len(codeList))

            preDay = self.code_edit2.text()

            for x in range(4):
                code = codeList[x]
                print(code)
                # SetInputValue
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", "20220214")
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")
                # CommRqData
                self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_pre_req", "opt10086", 0,
                                        "0102")
                time.sleep(0.3)

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
                print("이벤트 발생")
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
                self.receive_trCondition_loop.exit()
                time.sleep(0.3)

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

                print(abs(int(todayHigh)) - abs(int(self.preHigh)))
                print(abs(int(todayLow)) - abs(int(self.preLow)))
                print(abs(int(todayEnd)) - abs(int(self.preEnd)))
                print(abs(int(todayDay)) - abs(int(self.preDay)))
                self.receive_trCondition_loop.exit()



        except Exception as e:
            print(e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    app.exec_()
