import sys
import os
import cx_Oracle
import datetime
import time

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QAxContainer import *

# UI파일 연결
# UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("untitled.ui")[0]

# 오라클 DB 환경변수 등록
LOCATION = r"C:\instantclient_19_9"  # 32bit 버전 써야 한다
os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]

# 오라클 한글 지원
os.putenv('NLS_LANG', '.UTF8')


class startThread(QThread):
    # 쓰레드의 커스텀 이벤트
    # 데이터 전달 시 형을 명시해야 함
    threadEvent = pyqtSignal(str)
    set_accnt_event = pyqtSignal()

    def __init__(self, parent=None):  # 외부 클래스의 메소드를 사용하기 위해 파라미터로 outer_instance를 받는다
        # super().__init__()
        super(self.__class__, self).__init__(parent)
        self.g_is_thread = 0

    def run(self):
        # threadEvent 이벤트 발생
        # 파라미터 전달 기능(객체도 가능)
        self.threadEvent.emit('자동매매가 시작되었습니다.')

        set_tb_accnt_flag = 0  # 1이면 호출 완료

        while True:
            cur_tm = datetime.datetime.now()  # 현재시간 조회
            pre_tm = cur_tm.replace(hour=8, minute=30, second=1)
            open_tm = cur_tm.replace(hour=9, minute=0, second=1)
            close_tm = cur_tm.replace(hour=15, minute=30, second=1)
            if cur_tm - pre_tm >= datetime.timedelta(seconds=1):  # 8시 30분 이후라면
                # 계좌조회, 계좌정보 조회, 보유종목 매도주문 수행
                print('주문 중')
                # 계좌 조회, 계좌정보 조회, 보유종목 매도주문 수행
                if set_tb_accnt_flag == 0:
                    set_tb_accnt_flag = 1
                    self.set_accnt_event.emit()  # 계좌정보 요청 이벤트
            if cur_tm - open_tm >= datetime.timedelta(seconds=1):
                while True:
                    print('장 시작')
                    cur_tm = datetime.datetime.now()  # 현재시간 조회
                    if cur_tm - close_tm >= datetime.timedelta(seconds=1):  # 15시 30분 이후
                        print('장 종료')
                        break
                    # 장 운영시간 중이므로 매수나 매도주문
                    time.sleep(1)
            time.sleep(1)
            # break


# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.g_scr_no = 0  # Open API 요청번호
        self.g_user_id = None
        self.g_accnt_no = None

        # 키움증권 클래스를 사용하기 위해 인스턴스 생성(ProgID를 사용)
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # 이벤트(oneventconnect)와 이벤트 처리 메소드 연결
        # self.kiwoom.OnEventConnect.connect(self.event_connect)

        self.statusBar.showMessage('Ready')

        login_act = QAction('로그인', self)
        login_act.triggered.connect(self.menu_login_act)  # 로그인 메소드 호출
        self.menu_login.addAction(login_act)

        logout_act = QAction('로그아웃', self)
        logout_act.triggered.connect(self.menu_logout_act)  # 로그아웃 메소드 호출
        self.menu_logout.addAction(logout_act)

        # 콤보박스에서 선택한 계좌번호를 현재 계좌번호로 설정
        self.comboBox.currentIndexChanged.connect(self.comboBoxFunction)

        # 조회 버튼
        self.pushButton.clicked.connect(self.pushbutton_clicked)

        # 삽입 버튼
        self.pushButton_2.clicked.connect(self.pushbutton_2_clicked)

        # 수정 버튼
        self.pushButton_3.clicked.connect(self.pushbutton_3_clicked)

        # 삭제 버튼
        self.pushButton_4.clicked.connect(self.pushbutton_4_clicked)

        # 자동매매 시작
        self.pushButton_5.clicked.connect(self.pushbutton_5_clicked)

        # 자동매매 중지
        self.pushButton_6.clicked.connect(self.pushbutton_6_clicked)

        # 체크박스 생성
        self.insertTable()

        # 스레드 관련
        self.m_is_thread = 0  # 0이면 스레드 미생성, 1이면 스레드 생성
        # self.start_thread = WindowClass.startThread(self)

        # 쓰레드 인스턴스 생성
        self.start_thread = startThread()

        # 쓰레드 이벤트 연결
        self.start_thread.threadEvent.connect(self.threadEventHandler)
        self.start_thread.set_accnt_event.connect(self.set_tb_accnt)

        # self.thread1 = QThread()  # 스레드 생성. 파라미터 self???
        # self.start_thread.moveToThread(self.thread1)  # 만들어둔 쓰레드에 넣는다

        # 키움증권 open api 응답 대기 이벤트
        self.kiwoom.OnReceiveTrData.connect(self.axKHOpenAPI1_OnReceiveTrData)
        self.kiwoom.OnReceiveMsg.connect(self.axKHOpenAPI1_OnReceiveMsg)
        self.kiwoom.OnReceiveChejanData.connect(self.axKHOpenAPI1_OnReceiveChejanData)

        # 매수가능금액 데이터의 수신을 요청하고 수신 요청이 정상적으로 응답되는지 확인
        self.g_flag_1 = 0  # 1이면 요청에 대한 응답 완료

        # 데이터 수신을 요청할 때 전달할 요청명을 저장
        self.g_rqname = None

        # 매수가능액 저장
        self.g_ord_amt_possible = 0  # 총 매수가능 금액

    # 쓰레드 이벤트 핸들러
    # 장식자(데코레이터)에 파라미터 자료형을 명시
    @pyqtSlot(str)
    def threadEventHandler(self, s):
        self.write_msg_log(s)

    # 체크박스 생성 메소드
    def insertTable(self):
        self.checkBoxList = []
        for i in range(self.tableWidget.rowCount()):
            ckbox = QCheckBox()
            self.checkBoxList.append(ckbox)

        for i in range(self.tableWidget.rowCount()):
            cellWidget = QWidget()
            layoutCB = QHBoxLayout(cellWidget)
            layoutCB.addWidget(self.checkBoxList[i])
            layoutCB.setAlignment(Qt.AlignCenter)
            layoutCB.setContentsMargins(0, 0, 0, 0)
            cellWidget.setLayout(layoutCB)

            self.tableWidget.setCellWidget(i, 10, cellWidget)

    # 이벤트 처리 함수
    def event_connect(self, err_code):
        if err_code == 0:
            self.write_msg_log("로그인 성공")

    # 현재시각 가져오기
    def get_cur_tm(self):
        time = QTime.currentTime()
        answer = time.toString('hhmmss')

        return answer

    # 종목명 가져오기. 종목코드(문자열)를 매개변수로 받음
    def get_jongmok_nm(self, i_jongmok_cd):
        jongmok_nm = self.kiwoom.dynamicCall("GetMasterCodeName(%s)" % i_jongmok_cd)  # 종목명 가져오기

        return jongmok_nm

    # 오라클 접속 연결 메소드
    def connect_db(self):
        dsn = cx_Oracle.makedsn("localhost", 1521, "xe")
        db = cx_Oracle.connect("ats", "1234", "localhost/1521")

        return db

    # 로그 출력 메소드
    def write_msg_log(self, text):
        now = QDate.currentDate()
        datetime = QDateTime.currentDateTime()
        time = QTime.currentTime()
        cur_time = datetime.toString()
        cur_dt = now.toString(Qt.ISODate)
        cur_tm = time.toString('hh.mm.ss')
        cur_dtm = datetime.toString(Qt.DefaultLocaleShortDate)

        self.textEdit.append(cur_dtm + " " + text)

    # 오류 출력 메소드
    def write_err_log(self, text):
        now = QDate.currentDate()
        datetime = QDateTime.currentDateTime()
        time = QTime.currentTime()
        cur_time = datetime.toString()
        cur_dt = now.toString(Qt.ISODate)
        cur_tm = time.toString('hh.mm.ss')
        cur_dtm = datetime.toString(Qt.DefaultLocaleShortDate)

        self.textEdit_2.append(cur_dtm + " " + text)

    # 지연 메소드 time.sleep() 메소드

    # 요청번호 부여 메소드
    def get_scr_no(self):
        if self.g_scr_no < 9999:
            self.g_scr_no += 1
        else:
            self.g_scr_no = 1000

        return str(self.g_scr_no)

    # 로그인 버튼 액션 메소드
    def menu_login_act(self):
        ret = self.kiwoom.dynamicCall("CommConnect()")  # 로그인창 띄우기, OnEventConnect 이벤트 발생시킴
        # 로그인시 OnEventConnect 이벤트 발생, 계좌정보 가져오기 메소드를 실행
        self.kiwoom.OnEventConnect.connect(self.menu_login_status)

    # 로그아웃 버튼 액션 메소드
    def menu_logout_act(self):
        self.kiwoom.dynamicCall("CommTerminate()")  # 로그아웃
        self.statusBar.showMessage('로그아웃')

    # 계좌정보 가져오기 메소드
    def menu_login_status(self):
        accno = None  # 증권계좌번호
        accno_cnt = None  # 소유한 증권계좌번호의 수
        accno_arr = None  # N개의 증권계좌번호를 저장할 배열
        if self.kiwoom.dynamicCall("GetConnectState()") == 1:  # 로그인 성공
            self.g_user_id = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["USER_ID"])  # 반드시 리스트 형태로 값 전달하기
            accno = self.kiwoom.dynamicCall("GetLoginInfo(QString)", ["ACCNO"]).strip()  # 전체 계좌를 반환, 구분은 ';'로 돼있음,
            # 문자열 양 끝의 공백 제거
            accno_arr = accno.split(';')
            del accno_arr[len(accno_arr) - 1]  # 계좌리스트 마지막 공백 삭제

            self.lineEdit.setText(self.g_user_id)  # ID입력

            self.comboBox.clear()  # 콤보박스 초기화
            self.comboBox.addItems(accno_arr)  # 계좌 입력
            self.comboBox.setCurrentIndex(0)  # 첫번째 계좌번호를 콤보박스 초기 선택으로 설정
            self.g_accnt_no = self.comboBox.currentText()  # 설정된 계좌번호를 필드에 저장
        else:
            self.statusBar.showMessage('로그인 필요')

    # 콤보박스 선택시 수행
    def comboBoxFunction(self):
        self.g_accnt_no = str(self.comboBox.currentText())
        self.write_msg_log("사용할 증권계좌번호는 %s 입니다." % self.g_accnt_no)

    # dictinary 형태로 결과값 같기
    def makeDictFactory(self, cursor):
        columnNames = [d[0] for d in cursor.description]

        def createRow(*args):
            return dict(zip(columnNames, args))

        return createRow

    # 조회버튼 메소드
    def pushbutton_clicked(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        cur.execute("select * from TB_TRD_JONGMOK")

        cur.rowfactory = self.makeDictFactory(cur)

        rows = cur.fetchall()

        # 초기화
        row_num = 0
        cul_num = 1
        self.tableWidget.clearContents()  # 내용만 초기화
        self.insertTable()

        for row in rows:
            jongmok_cd = row['JONGMOK_CD']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(jongmok_cd))
            cul_num += 1
            jongmok_nm = row['JONGMOK_NM']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(jongmok_nm))
            cul_num += 1
            priority = row['PRIORITY']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(str(priority)))
            cul_num += 1
            buy_amt = row['BUY_AMT']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(str(buy_amt)))
            cul_num += 1
            buy_price = row['BUY_PRICE']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(str(buy_price)))
            cul_num += 1
            target_price = row['TARGET_PRICE']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(str(target_price)))
            cul_num += 1
            cut_loss_price = row['CUT_LOSS_PRICE']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(str(cut_loss_price)))
            cul_num += 1
            buy_trd_yn = row['BUY_TRD_YN']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(buy_trd_yn))
            cul_num += 1
            sell_trd_yn = row['SELL_TRD_YN']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(sell_trd_yn))
            cul_num = 1
            row_num += 1

        self.write_msg_log('TB_TRD_JONGMOK 테이블이 조회되었습니다')
        cur.close()
        conn.close()

    # 삽입버튼 메소드. 체크된 항목을 삽입
    def pushbutton_2_clicked(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        # print(self.tableWidget.rowCount())
        for row in range(0, self.tableWidget.rowCount()):
            print(row)
            if self.tableWidget.cellWidget(row, 10).findChild(type(QCheckBox())).isChecked():
                # if self.tableWidget.item(row, 10).checkState() == Qt.Checked:  # cellWidget에 체크박스를 넣지 않고 테이블에 바로 넣었을 경우
                user_id = self.g_user_id
                jongmok_cd = str(self.tableWidget.item(row, 1).text())
                jongmok_nm = str(self.tableWidget.item(row, 2).text())
                priority = int(self.tableWidget.item(row, 3).text())
                buy_amt = int(self.tableWidget.item(row, 4).text())
                buy_price = int(self.tableWidget.item(row, 5).text())
                target_price = int(self.tableWidget.item(row, 6).text())
                cut_loss_price = int(self.tableWidget.item(row, 7).text())
                buy_trd_yn = str(self.tableWidget.item(row, 8).text())
                sell_trd_yn = str(self.tableWidget.item(row, 9).text())

                # print(self.g_user_id)
                # print(jongmok_cd)
                # print(jongmok_nm)
                # print(priority)
                # print(buy_amt)
                # print(buy_price)
                # print(target_price)
                # print(cut_loss_price)
                # print(buy_trd_yn)
                # print(buy_trd_yn)
                # print(buy_trd_yn)
                # print(sell_trd_yn)

                try:
                    # sql_insert = f"insert into TB_TRD_JONGMOK values('{user_id}', '{jongmok_cd}', '{jongmok_nm}', {priority}, {buy_amt}, {buy_price}, {target_price}, {cut_loss_price}, '{buy_trd_yn}', '{sell_trd_yn}', '{user_id}', SYSDATE) "
                    sql_insert = 'insert into TB_TRD_JONGMOK VALUES(:USER_ID, :JONGMOK_CD, :JONGMOK_NM, :PRIORITY, ' \
                                 ':BUY_AMT, :BUY_PRICE, :TARGET_PRICE, :CUT_LOSS_PRICE, :BUY_TRD_YN, :SELL_TRD_YN, ' \
                                 ':INST_ID, :INST_DIM, NULL, NULL) '
                    cur.execute(sql_insert, USER_ID=user_id, JONGMOK_CD=jongmok_cd, JONGMOK_NM=jongmok_nm,
                                PRIORITY=priority, BUY_AMT=buy_amt, BUY_PRICE=buy_price, TARGET_PRICE=target_price,
                                CUT_LOSS_PRICE=cut_loss_price, BUY_TRD_YN=buy_trd_yn, SELL_TRD_YN=sell_trd_yn,
                                INST_ID=user_id, INST_DIM=datetime.datetime.now())
                    conn.commit()
                    self.write_msg_log('TB_TRD_JONGMOK 테이블이 변경되었습니다')
                except Exception as ex:
                    self.write_err_log("insert TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")
                    print("insert TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 수정버튼 메소드. 체크된 항목을 수정
    def pushbutton_3_clicked(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        # print(self.tableWidget.rowCount())
        for row in range(0, self.tableWidget.rowCount()):
            print(row)
            if self.tableWidget.cellWidget(row, 10).findChild(type(QCheckBox())).isChecked():
                # if self.tableWidget.item(row, 10).checkState() == Qt.Checked:  # cellWidget에 체크박스를 넣지 않고 테이블에 바로 넣었을 경우
                user_id = self.g_user_id
                jongmok_cd = str(self.tableWidget.item(row, 1).text())
                jongmok_nm = str(self.tableWidget.item(row, 2).text())
                priority = int(self.tableWidget.item(row, 3).text())
                buy_amt = int(self.tableWidget.item(row, 4).text())
                buy_price = int(self.tableWidget.item(row, 5).text())
                target_price = int(self.tableWidget.item(row, 6).text())
                cut_loss_price = int(self.tableWidget.item(row, 7).text())
                buy_trd_yn = str(self.tableWidget.item(row, 8).text())
                sell_trd_yn = str(self.tableWidget.item(row, 9).text())

                try:
                    sql_insert = 'UPDATE TB_TRD_JONGMOK SET JONGMOK_NM=:JONGMOK_NM, PRIORITY=:PRIORITY, ' \
                                 'BUY_AMT=:BUY_AMT, BUY_PRICE=:BUY_PRICE, TARGET_PRICE=:TARGET_PRICE, ' \
                                 'CUT_LOSS_PRICE=:CUT_LOSS_PRICE, BUY_TRD_YN=:BUY_TRD_YN, SELL_TRD_YN=:SELL_TRD_YN, ' \
                                 'UPDT_ID=:UPDT_ID, UPDT_DIM=:UPDT_DIM WHERE JONGMOK_CD = :jongmok_cd AND USER_ID = ' \
                                 ':user_id '
                    cur.execute(sql_insert, JONGMOK_NM=jongmok_nm,
                                PRIORITY=priority, BUY_AMT=buy_amt, BUY_PRICE=buy_price, TARGET_PRICE=target_price,
                                CUT_LOSS_PRICE=cut_loss_price, BUY_TRD_YN=buy_trd_yn, SELL_TRD_YN=sell_trd_yn,
                                UPDT_ID=user_id, UPDT_DIM=datetime.datetime.now(), jongmok_cd=jongmok_cd,
                                user_id=user_id)
                    conn.commit()
                    self.write_msg_log('TB_TRD_JONGMOK 테이블이 수정되었습니다')
                except Exception as ex:
                    self.write_err_log("UPDATE TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")
                    print("UPDATE TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 삭제버튼 메소드. 체크된 항목을 삭제
    def pushbutton_4_clicked(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        # print(self.tableWidget.rowCount())
        jongmok_cd = None
        for row in range(0, self.tableWidget.rowCount()):
            print(row)
            if self.tableWidget.cellWidget(row, 10).findChild(type(QCheckBox())).isChecked():
                # if self.tableWidget.item(row, 10).checkState() == Qt.Checked:  # cellWidget에 체크박스를 넣지 않고 테이블에 바로 넣었을 경우
                user_id = self.g_user_id
                jongmok_cd = str(self.tableWidget.item(row, 1).text())

                try:
                    sql_insert = 'DELETE FROM TB_TRD_JONGMOK WHERE JONGMOK_CD=:jongmok_cd AND USER_ID=:g_user_id'
                    cur.execute(sql_insert, jongmok_cd=jongmok_cd, g_user_id=user_id)
                    conn.commit()
                    self.write_msg_log('종목코드 : [' + jongmok_cd + ']가 삭제되었습니다.')
                except Exception as ex:
                    self.write_err_log("DELETE TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")
                    print("DELETE TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    def pushbutton_5_clicked(self):
        if self.m_is_thread == 1:  # 스레드가 이미 생성된 상태라면
            self.write_err_log('Auto Trading이 이미 시작되었습니다')
            return  # 메소드 종료
        self.m_is_thread = 1  # 스레드 생성으로 값 설정

        self.start_thread.start()  # 스레드 시작
        # self.start_thread.startWork()

    def pushbutton_6_clicked(self):
        self.write_msg_log('자동매매 중지 시작')
        try:
            self.start_thread.terminate()  # 현재 돌아가는 쓰레드 종료
            self.start_thread.wait()  # 새롭게 쓰레드를 대기시킴
        except Exception as ex:
            self.write_err_log('자동매매 중지 ex.Message : ' + str(ex))

        self.m_is_thread = 0

        self.write_msg_log('자동매매 중지 완료')

    # 투자정보 요청 이벤트 메소드
    # def axKHOpenAPI1_OnReceiveTrData(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1,
    #                                  msg2):
    def axKHOpenAPI1_OnReceiveTrData(self, screen_no, rqname, trcode, recordname, prev_next):
        if self.g_rqname == rqname:  # 요청한 요청명과 Open API로부터 응답받은 요청명이 같다면
            pass
        else:
            self.write_err_log("요청한 TR : [" + self.g_rqname + "]")
            self.write_err_log("응답받은 TR : [" + rqname + "]")

            if self.g_rqname == '증거금세부내역조회요청':
                self.g_flag_1 = 1  # 요청하는 쪽에서 무한루프에 빠지지 않게 방지
            return

        if rqname == '증거금세부내역조회요청':
            temp = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
                                                              trcode, '', rqname, 0,
                                                              '100주문가능금액') # 주문가능금액 저장
            self.g_ord_amt_possible = int(temp.strip())
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", screen_no)
            self.g_flag_1 = 1

    # 주식주문 요청 이벤트 메소드
    def axKHOpenAPI1_OnReceiveMsg(self, sender, e):
        a = None

    # 주식주문 내역, 체결내역 이벤트 메소드
    def axKHOpenAPI1_OnReceiveChejanData(self, sender, e):
        a = None

    # 매수가능금액 요청
    @pyqtSlot()
    def set_tb_accnt(self):
        for_cnt = 0
        for_flag = 0

        self.write_msg_log('TB_ACCNT 테이블 세팅 시작')
        self.g_ord_amt_possible = 0  # 매수가능금액

        for_flag = 0
        while True:
            self.kiwoom.dynamicCall("SetInputValue(QString, QString )", "계좌번호", self.g_accnt_no)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString )", "비밀번호", "")

            self.g_rqname = ""
            self.g_rqname = '증거금세부내역조회요청'  # 요청명 정의
            self.g_flag_1 = 0  # 요청 중

            scr_no = None  # 화면번호를 담을 변수 선언
            scr_no = self.get_scr_no()  # 화면번호 채번
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "증거금세부내역조회요청", "opw00013", 0,
                                    scr_no)  # 데이터 요청

            for_cnt = 0
            while True:  # 요청 후 대기 시작
                if self.g_flag_1 == 1:  # 응답 완료
                    time.sleep(1)
                    self.kiwoom.dynamicCall("DisconnectRealData(QString)", scr_no)
                    for_flag = 1
                    break
                else:
                    self.write_msg_log('증거금 세부내역 조회요청 완료 대기 중')
                    time.sleep(1)
                    for_cnt += 1
                    if for_cnt == 1:  # 한 번이라도 실패하면 무한루프 종료(증권계좌 비밀번호 오류 방지)
                        for_flag = 0
                        break
                    else:
                        continue
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", scr_no)
            if for_flag == 1:  # 요청에 대한 응답을 받았으므로 무한루프 종료
                break
            elif for_flag == 0:  # 요청에 대한 응답을 받지 못해도 비밀번호 5회 요류 방지를 위해 무한루프에서 빠져나옴
                time.sleep(1)
                break
            time.sleep(1)

        self.write_msg_log('주문가능금액 : [' + str(self.g_ord_amt_possible) + ']')
        self.merge_tb_accnt(self.g_ord_amt_possible)

    def merge_tb_accnt(self, g_ord_amt_possible):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()
        print(now.toString('ddMMyyyy'))


        try:
            # sql_insert = "MERGE INTO TB_ACCNT a USING(SELECT NVL(MAX(USER_ID), ' ') USER_ID, NVL(MAX(ACCNT_NO), ' ') ACCNT_NO, NVL(MAX(REF_DT), ' ') REF_DT FROM TB_ACCNT WHERE USER_ID = :g_user_id AND ACCNT_NO = :g_accnt_no AND REF_DT = :datecreate) b ON (a.USER_ID = b.USER_ID AND a.ACCNT_NO = b.ACCNT_NO AND a.REF_DT = b.REF_DT) WHEN MATCHED THEN UPDATE SET ORD_POSSIBLE_AMT = :g_ord_amt_possible, UPDT_DTM = :sysdate, UPDT_ID = 'ats' WHEN NOT MATCHED THEN INSERT (a.USER_ID, a.ACCNT_NO, a.REF_DT, a.ORD_POSSIBLE_AMT, a.INST_DTM, a.INST_ID) VALUES(:g_user_id, :g_accnt_no, :datecreate, :g_ord_amt_possible, :sysdate, 'ats')"
            sql_insert = "MERGE INTO TB_ACCNT a USING(SELECT NVL(MAX(USER_ID), ' ') USER_ID, NVL(MAX(ACCNT_NO), ' ') ACCNT_NO, NVL(MAX(REF_DT), ' ') REF_DT FROM TB_ACCNT WHERE USER_ID = :1 AND ACCNT_NO = :2 AND REF_DT = :3) b ON (a.USER_ID = b.USER_ID AND a.ACCNT_NO = b.ACCNT_NO AND a.REF_DT = b.REF_DT) WHEN MATCHED THEN UPDATE SET ORD_POSSIBLE_AMT = :4, UPDT_DTM = :5, UPDT_ID = 'ats' WHEN NOT MATCHED THEN INSERT (a.USER_ID, a.ACCNT_NO, a.REF_DT, a.ORD_POSSIBLE_AMT, a.INST_DTM, a.INST_ID) VALUES(:6, :7, :8, :9, :10, 'ats')"

            cur.execute(sql_insert, (self.g_user_id,
                        self.g_accnt_no, now.toString('yyyyMMdd'),
                        self.g_ord_amt_possible, datetime.datetime.now(),
                        self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd'),
                        self.g_ord_amt_possible, datetime.datetime.now()))
            conn.commit()
            self.write_msg_log('TB_ACCNT 테이블이 수정되었습니다')
        except Exception as ex:
            self.write_err_log("MERGE_TB_ACCNT ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()


if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
