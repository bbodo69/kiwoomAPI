import os
import sys
import time
import ClassMoeum
import datetime
import exchange_calendars as ecals
import math
import requests
import numpy
from PyQt5.QtCore import QEventLoop
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
from twilio.rest import Client

class MyWindow(QMainWindow):
    cl = ClassMoeum

    def __init__(self):
        super().__init__()

        self.StartRate = []
        self.HighRate = []
        self.LowRate = []
        self.EndRate = []

        # Kiwoom Login
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        #self.kiwoom.dynamicCall("CommConnect()")
        self.kiwoom.CommConnect()

        # Open API+ Event

        # 이벤트 루프 선언
        self.opt10086_req_loop = QEventLoop()
        self.event_connect_loop = QEventLoop()
        self.receive_trCondition_loop = QEventLoop()
        # 클래스로 접근법
        # self.kiwoom.OnEventConnect.connect(self.cl.event_connect)

        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)
        self.kiwoom.OnReceiveTrCondition.connect(self.receive_trCondition)
        self.kiwoom.OnReceiveMsg.connect(self.receive_msg)
        self.kiwoom.OnReceiveChejanData.connect(self.receive_Chejan)

        self.setWindowTitle("PyStock")
        self.setGeometry(600, 600, 600, 640)

        label = QLabel('입력란: ', self)
        label.move(20, 20)

        self.code_edit1 = QLineEdit(self)
        self.code_edit1.setGeometry(10, 10, 60, 30)
        self.code_edit1.move(90, 20)

        self.code_edit2 = QLineEdit(self)
        self.code_edit2.setGeometry(10, 10, 60, 30)
        self.code_edit2.move(160, 20)

        self.code_edit3 = QLineEdit(self)
        self.code_edit3.setGeometry(10, 10, 60, 30)
        self.code_edit3.move(230, 20)

        self.code_edit4 = QLineEdit(self)
        self.code_edit4.setGeometry(10, 10, 60, 30)
        self.code_edit4.move(300, 20)

        self.text_edit = QTextEdit(self)
        self.text_edit.setGeometry(10, 180, 500, 180)
        self.text_edit.setEnabled(True)

        self.listWidget = QListWidget(self)
        self.listWidget.setGeometry(10, 370, 500, 240)

        btn1 = QPushButton("조회", self)
        btn1.move(10, 60)
        btn1.clicked.connect(self.btn1_clicked)

        btn2 = QPushButton("계좌조회", self)
        btn2.move(120, 60)
        btn2.clicked.connect(self.btn2_clicked)

        btn3 = QPushButton("예수금상세현황요청", self)
        btn3.setGeometry(10, 10, 150, 30)
        btn3.move(230, 60)
        btn3.clicked.connect(self.btn3_clicked)

        btn4 = QPushButton("종목코드 얻기", self)
        btn4.move(10, 100)
        btn4.clicked.connect(self.btn4_clicked)

        btn5 = QPushButton("주식거래원요청", self)
        btn5.move(120, 100)
        btn5.clicked.connect(self.btn5_clicked)

        btn6 = QPushButton("조건검색 리스트", self)
        btn6.move(230, 100)
        btn6.clicked.connect(self.btn6_clicked)

        btn7 = QPushButton("조건검색", self)
        btn7.move(340, 100)
        btn7.clicked.connect(self.btn7_clicked)

        btn8 = QPushButton("조건검색중지", self)
        btn8.move(450, 100)
        btn8.clicked.connect(self.btn8_clicked)

        btn9 = QPushButton("주문", self)
        btn9.move(10, 140)
        btn9.clicked.connect(self.btn9_clicked)

        btn10 = QPushButton("일자비교", self)
        btn10.move(120, 140)
        btn10.clicked.connect(self.btn10_clicked)

        self.kiwoom.OnEventConnect.connect(self.event_connect)
        self.event_connect_loop.exec_()

    def event_connect(self, err_code):
        if err_code == 0:
            self.text_edit.append("로그인 성공")
            self.event_connect_loop.exit()

    def btn1_clicked(self):
        code = self.code_edit1.text()
        self.text_edit.setText("종목코드 : " + code)
        # self.text_edit.append("종목코드 : " + code)

        # SetInputValue
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")

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
            if data == "":
                self.text_edit.append("getConditionNameList(): 사용자 조건식이 없습니다.")

            conditionList = data.split(';')
            del conditionList[-1]

            self.listWidget.clear()
            self.listWidget.addItems(conditionList)

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
            self.receive_trCondition_loop.exec_()
            if isRequest == 0:
                print("조건검색 실패")
            if isRequest == 1:
                print("조건검색 접속 성공")

        except Exception as e:
            print(e)

    def btn8_clicked(self):
        try:
            screenNo = "0150"
            conditionName = self.code_edit1.text().split("^")[1]
            conditionIndex = self.code_edit1.text().split("^")[0]
            self.kiwoom.dynamicCall("SendConditionStop(QString, QString, int)",
                                     screenNo, conditionName, conditionIndex)
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
        except Exception as e:
            print(e)

    def btn10_clicked(self):
    # 주문
        try:
            XKRX = ecals.get_calendar("XKRX")  # 한국 코드

            # 라인 메세지 API
            TARGET_URL = 'https://notify-api.line.me/api/notify'
            TOKEN = 'HyzsVdhFD7USNTk4YsB21GXZvtPssAzPER9kTb0j7Xw'  # 발급받은 토큰
            headers = {'Authorization': 'Bearer ' + TOKEN}

            count = 0
            targetDay = 0
            progressCount = 0
            dayCount = 2

            for x in range(dayCount-1):

                self.StartRate.clear()
                self.HighRate.clear()
                self.LowRate.clear()
                self.EndRate.clear()

                # 17일 봉전 날짜
                while targetDay != x+2:
                    if XKRX.is_session(datetime.date.today() - datetime.timedelta(days=count)):
                        targetDay = targetDay + 1
                    count = count + 1

                preDay = str(datetime.date.today() - datetime.timedelta(days=count-1)).replace("-","")

                count = 0
                targetDay = 0

                # 18일 봉전 날짜
                while targetDay != x+1:
                    if XKRX.is_session(datetime.date.today() - datetime.timedelta(days=count)):
                        targetDay = targetDay + 1
                    count = count + 1

                todayDay = str(datetime.date.today() - datetime.timedelta(days=count-1)).replace("-","")

                # 일별주가요청 opt10086
                codeList = self.text_edit.toPlainText().split(',')

                print("종목수" + str(len(codeList)))

                if len(codeList) > math.floor(500/dayCount):
                    timeDelay = 3.6
                elif len(codeList) > math.floor(50/dayCount) and len(codeList) <= math.floor(500/dayCount):
                    timeDelay = 1.8
                elif len(codeList) <= math.floor(50/dayCount):
                    timeDelay = 0.2

                print("타임딜레이" + str(timeDelay))

                for code in codeList:
                    # 작업 진행 상태
                    progressCount = progressCount + 1
                    print(str(progressCount) + " / " + str(len(codeList) * (dayCount-1)))

                    # SetInputValue
                    code = code.replace(" ","")
                    self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
                    self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", preDay)
                    self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")
                    # CommRqData
                    self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_pre_req", "opt10086", 0,
                                            "0103")
                    self.opt10086_req_loop.exec()
                    time.sleep(timeDelay)

                    # SetInputValue
                    self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
                    self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", todayDay)
                    self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")
                    # CommRqData
                    self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_today_req", "opt10086",
                                            0,
                                            "0103")
                    self.opt10086_req_loop.exec()
                    time.sleep(timeDelay)

                # 앞 뒤 5% 데이터 리스트에서 삭제
                iCount = round(len(self.StartRate) * 0.05)

                self.StartRate.sort()
                self.HighRate.sort()
                self.LowRate.sort()
                self.EndRate.sort()

                if iCount >= 1:
                    for x in range(iCount):
                        self.StartRate.pop()
                        del self.StartRate[0]
                        self.HighRate.pop()
                        del self.HighRate[0]
                        self.LowRate.pop()
                        del self.LowRate[0]
                        self.EndRate.pop()
                        del self.EndRate[0]

                body_text_info = "\n" + "날짜 : " + todayDay + "\n" + \
                            "종목 갯수 : " + str(len(self.StartRate)) + "\n" + \
                            "시가 평균 : " + str(round(sum(self.StartRate)/len(self.StartRate),4)) + "\n" + \
                            "고가 평균 : " + str(round(sum(self.HighRate) / len(self.HighRate),4)) + "\n" + \
                            "저가 평균 : " + str(round(numpy.mean(self.LowRate), 4)) + "\n" + \
                            "종가 평균 : " + str(round(numpy.mean(self.EndRate), 4))  + "\n" + \
                            "종가 상위 5개 : " + "{{top5}}" + "\n" + \
                            "종가 하위 5개 : " + "{{bottom5}}"

                top5 = ""
                bottom5 = ""

                if iCount >= 1:
                    for x in range(iCount):
                        top5 += str(self.EndRate[len(self.EndRate) - (x+1)]) + ", "
                        bottom5 += str(self.EndRate[x]) + ", "

                body_text = {'message': body_text_info.replace("{{top5}}", top5).replace("{{bottom5}}", bottom5)

#                            "종가 상위 5개 : "
#                                        + str(self.EndRate[len(self.EndRate) - 1]) + ", "
#                                        + str(self.EndRate[len(self.EndRate) - 2]) + ", "
#                                        + str(self.EndRate[len(self.EndRate) - 3]) + ", "
#                                        + str(self.EndRate[len(self.EndRate) - 4]) + ", "
#                                        + str(self.EndRate[len(self.EndRate) - 5]) + ", " + "\n" + \
#                            "종가 하위 5개 : " + str(self.EndRate[0]) + ", "
#                                        + str(self.EndRate[1]) + ", "
#                                        + str(self.EndRate[2]) + ", "
#                                        + str(self.EndRate[3]) + ", "
#                                        + str(self.EndRate[4]) + ", "
                             }

                requests.post(TARGET_URL, headers=headers, data=body_text)
                print("날짜 : " + todayDay)
                print("시가 평균 : " + str(round(sum(self.StartRate)/len(self.StartRate),4)))
                print("고가 평균 : " + str(round(sum(self.HighRate) / len(self.HighRate),4)))
                print("저가 평균 : " + str(round(numpy.mean(self.LowRate),4)))
                print("종가 평균 : " + str(round(numpy.mean(self.EndRate),4)))

            #os.system('shutdown -s -t 100')
        except Exception as e:
            #os.system('shutdown -s -t 100')
            body_text = {'message': "에러 : " + str(e)}
            requests.post(TARGET_URL, headers=headers, data=body_text)
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
        print("1")
        try:
            if codes == "":
                print("검색결과 없음")
                return
            codeList = codes.split(';')
            del codeList[-1]
            self.text_edit.setText(', '.join(codeList))
            self.receive_trCondition_loop.exit()

        except Exception as e:
            print(e)

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
                self.preUpDown = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                              rqname, 0, "등락률")

                #print(self.preStart)
                #print(self.preHigh)
                #print(self.preLow)
                #print(self.preEnd)
                #print(self.preDay)
                self.opt10086_req_loop.exit()

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

                # 등락률 = (해당가격 - 전일종가)/전일종가

                self.StartRate.append(round((abs(int(todayStart)) - abs(int(self.preEnd)))/abs(int(self.preEnd)), 4))
                self.HighRate.append(round((abs(int(todayHigh)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                self.LowRate.append(round((abs(int(todayLow)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                self.EndRate.append(round((abs(int(todayEnd)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                #print(todayStart)
                #print(todayHigh)
                #print(todayLow)
                #print(todayEnd)
                #print(todayDay)

                #print(abs(int(todayHigh)) - abs(int(self.preHigh)))
                #print(abs(int(todayLow)) - abs(int(self.preLow)))
                #print(abs(int(todayEnd)) - abs(int(self.preEnd)))
                #print(abs(int(todayDay)) - abs(int(self.preDay)))
                self.opt10086_req_loop.exit()

        except Exception as e:
            print(e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
