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

class MyWindow():
    cl = ClassMoeum
    try:
        def __init__(self):
            super().__init__()

            self.StartRate = []
            self.HighRate = []
            self.LowRate = []
            self.EndRate = []

        # 각종 플래그 설정
            # 매수 진행여부 트리거
            self.moiTrade = False


            # Kiwoom Login
            self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
            self.kiwoom.CommConnect()

            # Open API+ Event

            # 이벤트 루프 선언
            self.opt10086_req_loop = QEventLoop()
            self.opt10001_req_loop = QEventLoop()
            self.event_connect_loop = QEventLoop()
            self.receive_trCondition_loop = QEventLoop()
            self.GetConditionLoad_loop = QEventLoop()
            self.receive_conditionVer_loop = QEventLoop()
            # 클래스로 접근법
            # self.kiwoom.OnEventConnect.connect(self.cl.event_connect)

            self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)
            self.kiwoom.OnReceiveTrCondition.connect(self.receive_trCondition)
            self.kiwoom.OnReceiveMsg.connect(self.receive_msg)
            self.kiwoom.OnReceiveChejanData.connect(self.receive_Chejan)
            self.kiwoom.OnReceiveConditionVer.connect(self.receive_conditionVer)

            self.kiwoom.OnEventConnect.connect(self.event_connect)
            self.event_connect_loop.exec_()

            # 계좌정보 불러오기
            account_num = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"])
            server_info = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")
            account_cnt = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "ACCOUNT_CNT")

            # 두번째 계좌정보 사용
            self.useAccountNum = str(account_num).split(";")[1]
            print("계좌정보 : " + str(self.useAccountNum))

        # # 조건검색식 불러오기
        #     # isLoad = self.kiwoom.dynamicCall("GetConditionLoad()")
        #     print("조건검색식 불러오기")
        #     isLoad = self.kiwoom.GetConditionLoad()
        #     self.receive_conditionVer_loop.exec_()
        #
        #     if isLoad == 0:
        #         print("조건식 요청 실패")
        #     data = self.kiwoom.dynamicCall("GetConditionNameList()")
        #     if data == "":
        #         print("getConditionNameList(): 사용자 조건식이 없습니다.")
        #
        #     self.conditionList = data.split(';')
        #     del self.conditionList[-1]
        #     print("조건식 리스트 검색 : " + str(self.conditionList))
        #     self.condition = self.conditionList[-2]
        #
        #     # 조건식 검색 종목명 추출
        #     screenNo = "0150"
        #     print("사용 조건식 : " + str(self.condition))
        #     conditionName = self.condition.split("^")[1]
        #     conditionIndex = int(self.condition.split("^")[0])
        #     print("조건식 이름 : " + conditionName + " " + str(type(conditionName)))
        #     print("조건식 번호 : " + str(conditionIndex) + " " + str(type(conditionIndex)))
        #     isRealTime = 1
        #     isRequest = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)",
        #                                         screenNo, conditionName, conditionIndex, isRealTime)
        #     self.receive_trCondition_loop.exec_()
        #     print(isRequest)
        #     if isRequest == 0:
        #         print("조건검색 실패")
        #     if isRequest == 1:
        #         print("조건검색 접속 성공")

            # 검색 종목명 가격 비교

            #self.btn10_clicked()

        # 조건검색식 불러오기
            # isLoad = self.kiwoom.dynamicCall("GetConditionLoad()")
            print("조건검색식 불러오기")
            isLoad = self.kiwoom.GetConditionLoad()
            self.receive_conditionVer_loop.exec_()

            if isLoad == 0:
                print("조건식 요청 실패")
            data = self.kiwoom.dynamicCall("GetConditionNameList()")
            if data == "":
                print("getConditionNameList(): 사용자 조건식이 없습니다.")

            self.conditionList = data.split(';')
            del self.conditionList[-1]
            print("조건식 리스트 : " + str(self.conditionList))
            
            # 조건식 선택
            self.condition = self.conditionList[-1]

            # 조건식 검색 종목명 추출
            screenNo = "0150"
            print("사용 조건식 : " + str(self.condition))
            conditionName = self.condition.split("^")[1]
            conditionIndex = int(self.condition.split("^")[0])
            print("조건식 이름 : " + conditionName + " " + str(type(conditionName)))
            print("조건식 번호 : " + str(conditionIndex) + " " + str(type(conditionIndex)))
            isRealTime = 1            
            # self.codeList 에 코드리스트 저장
            isRequest = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)",
                                                screenNo, conditionName, conditionIndex, isRealTime)
            self.receive_trCondition_loop.exec_()
            print(isRequest)
            if isRequest == 0:
                print("조건검색 실패")
            if isRequest == 1:
                print("조건검색 접속 성공")

            # 검색 종목명 가격 비교
            #self.btn10_clicked()

            # 매수 진행 (매개변수 코드리스트 필요)
            self.buying(self.codeList)
            self.SendLineMessage()
            # self.powreOff()

    except Exception as e:
        print(e)

            # # 조건식 결과 종목 매수 진행
            #
            # orderAmount = '1'
            # code = '005110'
            # orderPrice = '701'
            # # 사용자구분명, 화면번호, 계좌번호 10자리, 주문유형 1~7, 종목코드(6자리), 주문수량, 주문가격, 거래구분 00~81, 원주문번호
            #
            # isRequest = self.kiwoom.SendOrder("SendOrder_req", "0211", "5061614811", 1, code,
            #                                   int(orderAmount), int(orderPrice), "00", "")
            #
            # if isRequest == 0:
            #     print("매수주문 성공")
            # else:
            #     print("매수주문 실패 : " + str(isRequest))

    # 로그인
    def event_connect(self, err_code):
        if err_code == 0:
            self.event_connect_loop.exit()

    def btn1_clicked(self):
        # self.text_edit.append("종목코드 : " + code)

        # SetInputValue
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", "")

        # CommRqData
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")

    def btn2_clicked(self):
        account_num = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"])
        server_info = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "GetServerGubun")
        account_cnt = self.kiwoom.dynamicCall("GetLoginInfo(QString)", "ACCOUNT_CNT")

    def btn3_clicked(self):
        #account_num = self.code_edit1.text()
        #self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "계좌번호", account_num)
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회구분", "2")

        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opw00001", "opw00001", 0, "0101")

    def btn6_clicked(self):
    # 조건식 리스트 불러오기
        try:
            #isLoad = self.kiwoom.dynamicCall("GetConditionLoad()")
            isLoad = self.kiwoom.GetConditionLoad()
            if isLoad == 0:
                print("조건식 요청 실패")

            data = self.kiwoom.dynamicCall("GetConditionNameList()")
            if data == "":
                time.sleep(5)
                print("getConditionNameList(): 사용자 조건식이 없습니다.")
                self.btn6_clicked()

            self.conditionList = data.split(';')
            del self.conditionList[-1]
            print("조건식 리스트 검색 : " + str(self.conditionList))

        except Exception as e:
            print(e)

    def btn7_clicked(self):
    # 조건식 검색
        try:
            screenNo = "0150"
            condition = self.conditionList[-1]
            print("조건식 검색 : " + str(condition))
            conditionName = condition.split("^")[1]
            conditionIndex = condition.split("^")[0]
            isRealTime = "1"
            isRequest = self.kiwoom.dynamicCall("SendCondition(QString, QString, int, int)",
                                     screenNo, conditionName, conditionIndex, isRealTime)
            print("test")
            self.receive_trCondition_loop.exec_()
            print("test1")
            if isRequest == 0:
                print("조건검색 실패")
            if isRequest == 1:
                print("조건검색 접속 성공")

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

                print("종목수" + str(len(self.codeList)))

                if len(self.codeList) > math.floor(250/dayCount):
                    timeDelay = 3.6
                elif len(self.codeList) > math.floor(25/dayCount) and len(self.codeList) <= math.floor(250/dayCount):
                    timeDelay = 1.8
                elif len(self.codeList) <= math.floor(25/dayCount):
                    timeDelay = 0.2

                timeDelay = 3.6

                print("타임딜레이" + str(timeDelay))

                for code in self.codeList:
                    # 작업 진행 상태
                    progressCount = progressCount + 1
                    print(str(progressCount) + " / " + str(len(self.codeList) * (dayCount-1)))

                    # SetInputValue
                    code = code.replace(" ","")
                    print(code)
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
    
    # 전원끄기
    def powreOff(self):
        os.system('shutdown -s -t 10')

    def buying(self, codeList):

        self.unit = ""
        for code in codeList:
            # SetInputValue
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)

            # CommRqData, 호가단위 확인 'self.unit' 에 호가단위 저장
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")
            time.sleep(3.6)
            self.opt10001_req_loop.exec_()
            print("호가단위 = " + str(self.unit))

            # 전일비 비교 데이터 추출
            # SetInputValue
            code = code.replace(" ", "")
            print(code)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", self.calculBusDay(1))
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")
            # CommRqData
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_pre_req", "opt10086", 0,
                                    "0103")
            time.sleep(3.6)
            self.opt10086_req_loop.exec()

            # 전일 종가 확인 self.todayEnd 에 저장
            # SetInputValue
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", code)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "조회일자", self.calculBusDay(0))
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "표시구분", "1")
            # CommRqData
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10086_today_req", "opt10086",
                                    0,
                                    "0103")
            time.sleep(3.6)
            self.opt10086_req_loop.exec()

            print("전일종가 : " + str(self.preEnd).strip())

            #주문 진행 , 종가 기준 1%, 호가기준 가격내림
            buyPrice = abs(int(self.preEnd)*0.99)

            if self.unit == 100:
                buyPrice = math.floor(buyPrice/100) * 100
            elif self.unit == 50:
                buyPrice = math.floor(buyPrice/50) * 50
            elif self.unit == 10:
                buyPrice = math.floor(buyPrice/10) * 10
            elif self.unit == 5:
                buyPrice = math.floor(buyPrice/5)*5
            elif self.unit == 1:
                buyPrice = math.floor(buyPrice)

            print("주문가 : " + str(buyPrice))
            print(self.EndRate)

            # 조건식 결과 종목 매수 진행
            if self.moiTrade == True:
                orderAmount = round(500000 / buyPrice)
                orderPrice = buyPrice

                # 사용자구분명, 화면번호, 계좌번호 10자리, 주문유형 1~7, 종목코드(6자리), 주문수량, 주문가격, 거래구분 00~81, 원주문번호
                accountNumber = self.useAccountNum
                self.SendOrderRequest = self.kiwoom.SendOrder("SendOrder_req", "0211", accountNumber, 1, code,
                                                   int(orderAmount), int(orderPrice), "00", "")

                self.SendOrderSucceseCount = 0
                if self.SendOrderRequest == 0:
                    print("매수주문 성공")
                    self.SendOrderSucceseCount += 1
                else:
                    print("매수주문 실패")




    # 라인 메세지 전송
    def SendLineMessage(self):

        # 라인 메세지 API
        TARGET_URL = 'https://notify-api.line.me/api/notify'
        TOKEN = 'HyzsVdhFD7USNTk4YsB21GXZvtPssAzPER9kTb0j7Xw'  # 발급받은 토큰
        headers = {'Authorization': 'Bearer ' + TOKEN}

        iCount = round(len(self.StartRate) * 0.05)

        body_text_info = "\n" + "검색식 : " + str(self.condition) + "\n" + \
                         "날짜 : " + self.calculBusDay(0) + "\n" + \
                         "종목 갯수 : " + str(len(self.StartRate)) + "\n" + \
                         "시가 평균 : " + str(round(sum(self.StartRate) / len(self.StartRate), 4)) + "\n" + \
                         "고가 평균 : " + str(round(sum(self.HighRate) / len(self.HighRate), 4)) + "\n" + \
                         "저가 평균 : " + str(round(numpy.mean(self.LowRate), 4)) + "\n" + \
                         "종가 평균 : " + str(round(numpy.mean(self.EndRate), 4)) + "\n" + \
                         "종가 상위 5% : " + "{{top5}}" + "\n" + \
                         "종가 하위 5% : " + "{{bottom5}}" + "\n" + \
                         "종가 1% 이상 개수 : " + "{{EndRate4}}" + "\n" + \
                         "종가 0~1% 개수 : " + "{{EndRate3}}" + "\n" + \
                         "종가 -1~0% 개수 : " + "{{EndRate2}}" + "\n" + \
                         "종가 -1% 이하 개수 : " + "{{EndRate1}}" + "\n" + \
                         "고가 1% 이상 개수 : " + "{{HighRate}}" + "\n" + \
                         "저가 -1% 개수 : " + "{{LowRate}}" + "\n" + \
                         "총 주문 / 성공 주문 : " + "{{total}} / {{SuccessOrder}}"
        top5 = ""
        bottom5 = ""

        self.EndRate.sort()

        # 종가 상위, 하위 5개 추출
        if iCount >= 1:

            for x in range(iCount):
                top5 += str(self.EndRate[len(self.EndRate) - (x + 1)]) + ", "
                bottom5 += str(self.EndRate[x]) + ", "

        # 종가 -1%, -1~0%, 0~1%, 1% 갯수 추출
        EndRate1 = 0
        EndRate2 = 0
        EndRate3 = 0
        EndRate4 = 0

        for x in self.EndRate:
            if x < -0.01:
                EndRate1 += 1
            elif x >= -0.01 and x < 0:
                EndRate2 += 1
            elif x >= 0 and x <= 0.01:
                EndRate3 += 1
            elif x > 0.01:
                EndRate4 += 1

        # 고가, 저가 1% 이상 갯수
        HighRate = 0
        LowRate = 0

        for x in self.HighRate:
            if x >= 0.01:
                HighRate += 1

        for x in self.LowRate:
            if x <= -0.01:
                LowRate += 1

        body_text = {'message': body_text_info.replace("{{top5}}", top5).replace("{{bottom5}}", bottom5)
            .replace("EndRate1", str(EndRate1)).replace("EndRate2", str(EndRate2)).
            replace("EndRate3", str(EndRate3)).replace("EndRate4", str(EndRate4)).replace("HighRate", str(HighRate))
            .replace("LowRate", str(LowRate)).replace("{{total}}", str(len(self.StartRate))).replace("{{SuccessOrder}}", str(self.SendOrderSucceseCount))}

        requests.post(TARGET_URL, headers=headers, data=body_text)

    # 엽업일전 날짜 계산 return 값 : day 영업일전 날짜
    def calculBusDay(self, day):

        XKRX = ecals.get_calendar("XKRX")  # 한국 코드

        # 라인 메세지 API
        TARGET_URL = 'https://notify-api.line.me/api/notify'
        TOKEN = 'HyzsVdhFD7USNTk4YsB21GXZvtPssAzPER9kTb0j7Xw'  # 발급받은 토큰
        headers = {'Authorization': 'Bearer ' + TOKEN}

        count = 0
        targetDay = 0
        dayCount = day + 2

        for x in range(dayCount - 1):

            while targetDay != x + 1:
                if XKRX.is_session(datetime.date.today() - datetime.timedelta(days=count)):
                    targetDay = targetDay + 1
                count = count + 1

            todayDay = str(datetime.date.today() - datetime.timedelta(days=count - 1)).replace("-", "")

        return todayDay

    def receive_conditionVer(self, sMsg):
        self.receive_conditionVer_loop.exit()
        print("sMsg = " + str(sMsg))

    def receive_Chejan(self, sGubun, nItemCnt,sFIdList):
        #체결구분. 접수와 체결시 '0'값, 국내주식 잔고변경은 '1'값, 파생잔고변경은 '4'

        # 라인 메세지 API
        TARGET_URL = 'https://notify-api.line.me/api/notify'
        TOKEN = 'HyzsVdhFD7USNTk4YsB21GXZvtPssAzPER9kTb0j7Xw'  # 발급받은 토큰
        headers = {'Authorization': 'Bearer ' + TOKEN}

        body_text = {'message': "체결완료"}
        requests.post(TARGET_URL, headers=headers, data=body_text)

    def receive_msg(self, screenNo, rqname, trcode, msg):
        try:
            print(msg)
        except Exception as e:
            print(e)


    def receive_trCondition(self, screenNo, codes, conditionName, conditionIndex, inquiry):
    # 조건검색 결과 리스트 불러오기
        print("조건검색 결과 리스트 불러오기")
        try:
            if codes == "":
                print("검색결과 없음")
                return
            self.codeList = codes.split(';')
            del self.codeList[-1]
            self.receive_trCondition_loop.exit()

        except Exception as e:
            print(e)

    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        try:
            print("요청이름 : " + rqname)

            # 호가단위 확인
            if rqname == "opt10001_req":
                name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "종목명")
                #volume = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                #                                 rqname, 0, "거래량")
                self.price = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "현재가")
                #BPS = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                #                                 rqname, 0, "BPS")
                #unit = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                #                                 rqname, 0, "액면가단위")

                #self.text_edit.append("종목명: " + name.strip())
                #self.text_edit.append("거래량: " + volume.strip())
                print("종목명: " + name.strip())
                print("현재가: " + self.price.strip())
                #self.text_edit.append("BPS: " + BPS.strip())
                self.opt10001_req_loop.exit()

                self.price = abs(int(self.price))
                if self.price >= 50000:
                    self.unit = 100
                # 10,000원 이상~50,000원 미만
                elif 10000 <= self.price < 50000:
                    self.unit = 50
                elif 5000 <= self.price < 10000:
                    self.unit = 10
                elif 1000 <= self.price < 5000:
                    self.unit = 5
                elif self.price < 1000:
                    self.unit = 1

            if rqname == "opt10002_req":
                name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "종목명")

                bound = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "등락율")

                #self.text_edit.append("종목명: " + name.strip())
                #self.text_edit.append("bound: " + bound.strip())

                #if price < BPS:
                    #self.text_edit.append("종목명: " + name.strip())
                    #self.text_edit.append("현재가: " + price.strip())

            if rqname == "opw00001":
                diposit = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "예수금")

                #self.text_edit.setText("예수금: " + diposit.strip())


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
                self.todayStart = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                               rqname, 0, "시가")
                self.todayHigh = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "고가")
                self.todayLow = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "저가")
                self.todayEnd = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                                 rqname, 0, "종가")
                self.todayDay = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "",
                                              rqname, 0, "전일비")

                # 등락률 = (해당가격 - 전일종가)/전일종가

                #print(todayStart)
                #print(todayHigh)
                #print(todayLow)
                #print(todayEnd)
                #print(todayDay)

                self.StartRate.append(round((abs(int(self.todayStart)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                self.HighRate.append(round((abs(int(self.todayHigh)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                self.LowRate.append(round((abs(int(self.todayLow)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                self.EndRate.append(round((abs(int(self.todayEnd)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                print("전일종가 : {0}, 금일종가 : {1}".format(str(self.preEnd).strip(), str(self.todayEnd).strip()))
                # print(abs(int(self.todayHigh)) - abs(int(self.preHigh)))
                # print(abs(int(self.todayLow)) - abs(int(self.preLow)))
                # print(abs(int(self.todayEnd)) - abs(int(self.preEnd)))
                # print(abs(int(self.todayDay)) - abs(int(self.preDay)))
                print(round((abs(int(self.todayEnd)) - abs(int(self.preEnd))) / abs(int(self.preEnd)), 4))
                self.opt10086_req_loop.exit()

        except Exception as e:
            print(e)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    app.exec_()
