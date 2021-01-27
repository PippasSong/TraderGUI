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


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


# UI파일 연결
# UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
# form_class = uic.loadUiType("untitled.ui")[0]
form = resource_path('untitled.ui')
form_class = uic.loadUiType(form)[0]

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
    set_accnt_info_event = pyqtSignal()
    ord_first_event = pyqtSignal()
    real_buy_ord_event = pyqtSignal()
    real_sell_ord_event = pyqtSignal()
    real_cut_loss_ord_event = pyqtSignal()

    def __init__(self, parent=None):  # 외부 클래스의 메소드를 사용하기 위해 파라미터로 outer_instance를 받는다
        # super().__init__()
        super(self.__class__, self).__init__(parent)
        self.g_is_thread = 0

    def run(self):
        # threadEvent 이벤트 발생
        # 파라미터 전달 기능(객체도 가능)
        self.threadEvent.emit('자동매매가 시작되었습니다.')

        set_tb_accnt_flag = 0  # 1이면 호출 완료
        set_tb_accnt_info_flag = 0  # 1이면 호출 완료
        sell_ord_first_flag = 0  # 1이면 호출 완료

        while True:
            cur_tm = datetime.datetime.now()  # 현재시간 조회
            pre_tm = cur_tm.replace(hour=8, minute=30, second=1)
            open_tm = cur_tm.replace(hour=9, minute=0, second=1)
            close_tm = cur_tm.replace(hour=15, minute=30, second=1)
            if cur_tm - pre_tm >= datetime.timedelta(seconds=1):  # 8시 30분 이후라면
                # 계좌조회, 계좌정보 조회, 보유종목 매도주문 수행
                if set_tb_accnt_flag == 0:
                    set_tb_accnt_flag = 1
                    self.set_accnt_event.emit()  # 계좌정보 요청 이벤트
                if set_tb_accnt_info_flag == 0:
                    self.set_accnt_info_event.emit()  # 계좌정보 조회 요청 이벤트
                    set_tb_accnt_info_flag = 1
                if sell_ord_first_flag == 0:
                    self.ord_first_event.emit()  # 보유종목 매도 이벤트
                    sell_ord_first_flag = 1

            if cur_tm - open_tm >= datetime.timedelta(seconds=1):  # 장 운영시간
                while True:
                    cur_tm = datetime.datetime.now()  # 현재시간 조회
                    if cur_tm - close_tm >= datetime.timedelta(seconds=1):  # 15시 30분 이후
                        break
                    # 장 운영시간 중이므로 매수나 매도주문
                    self.real_buy_ord_event.emit()  # 실시간 매수주문 이벤트
                    time.sleep(10)
                    self.real_sell_ord_event.emit()  # 실시간 매도주문 이벤트
                    time.sleep(0.2)
                    self.real_cut_loss_ord_event.emit()  # 실시간 손절주문 이벤트

            time.sleep(1)


# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.g_scr_no = 0  # Open API 요청번호
        self.g_user_id = None
        self.g_accnt_no = None
        self.g_cur_price = 0

        self.g_buy_hoga = 0  # 최우선 매수호가 저장 변수

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
        self.start_thread.set_accnt_info_event.connect(self.set_tb_accnt_info)
        self.start_thread.ord_first_event.connect(self.sell_ord_first)
        self.start_thread.real_buy_ord_event.connect(self.real_buy_ord)
        self.start_thread.real_sell_ord_event.connect(self.real_sell_ord)
        self.start_thread.real_cut_loss_ord_event.connect(self.real_cut_loss_ord)

        # self.thread1 = QThread()  # 스레드 생성. 파라미터 self???
        # self.start_thread.moveToThread(self.thread1)  # 만들어둔 쓰레드에 넣는다

        # 키움증권 open api 응답 대기 이벤트
        self.kiwoom.OnReceiveTrData.connect(self.axKHOpenAPI1_OnReceiveTrData)
        self.kiwoom.OnReceiveMsg.connect(self.axKHOpenAPI1_OnReceiveMsg)
        self.kiwoom.OnReceiveChejanData.connect(self.axKHOpenAPI1_OnReceiveChejanData)

        # 매수가능금액 데이터의 수신을 요청하고 수신 요청이 정상적으로 응답되는지 확인
        self.g_flag_1 = 0  # 1이면 요청에 대한 응답 완료
        # 계좌정보 데이터 수신 요청
        self.g_flag_2 = 0  # 1이면 요청에 대한 응답 완료
        self.g_is_next = 0  # 다음 조회 데이터가 있는지 확인. 한번에 50개의 주식만 조회할 수 있기 때문

        self.g_flag_3 = 0  # 매수주문 응답 플래그
        self.g_flag_4 = 0  # 매도주문 응답 플래그
        self.g_flag_5 = 0  # 매도취소주문 응답 플래그
        self.g_flag_6 = 0  # 현재가 조회 플래그 변수가 1이면 조회 완료
        self.g_flag_7 = 0  # 최우선 매수호가 플래스 변수가 1이면 조회 완료

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
        jongmok_nm = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", i_jongmok_cd)  # 종목명 가져오기

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


        cur.close()
        conn.close()

    # 수정버튼 메소드. 체크된 항목을 수정
    def pushbutton_3_clicked(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        # print(self.tableWidget.rowCount())
        for row in range(0, self.tableWidget.rowCount()):
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
                                 'UPDT_ID=:UPDT_ID, UPDT_DTM=:UPDT_DTM WHERE JONGMOK_CD = :jongmok_cd AND USER_ID = ' \
                                 ':user_id '
                    cur.execute(sql_insert, JONGMOK_NM=jongmok_nm,
                                PRIORITY=priority, BUY_AMT=buy_amt, BUY_PRICE=buy_price, TARGET_PRICE=target_price,
                                CUT_LOSS_PRICE=cut_loss_price, BUY_TRD_YN=buy_trd_yn, SELL_TRD_YN=sell_trd_yn,
                                UPDT_ID=user_id, UPDT_DTM=datetime.datetime.now(), jongmok_cd=jongmok_cd,
                                user_id=user_id)
                    conn.commit()
                    self.write_msg_log('TB_TRD_JONGMOK 테이블이 수정되었습니다')
                except Exception as ex:
                    self.write_err_log("UPDATE TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 삭제버튼 메소드. 체크된 항목을 삭제
    def pushbutton_4_clicked(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        # print(self.tableWidget.rowCount())
        jongmok_cd = None
        for row in range(0, self.tableWidget.rowCount()):
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

        cur.close()
        conn.close()

    def pushbutton_5_clicked(self):
        if self.m_is_thread == 1:  # 스레드가 이미 생성된 상태라면
            self.write_err_log('Auto Trading이 이미 시작되었습니다')
            return  # 메소드 종료
        self.m_is_thread = 1  # 스레드 생성으로 값 설정

        self.start_thread.start()  # 스레드 시작

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
            elif self.g_rqname == '계좌평가현황요청':
                self.g_flag_2 = 1  # 요청하는 쪽에서 무한루프에 빠지지 않게 방지
            elif self.g_rqname == '호가조회':
                self.g_flag_7 = 1
            elif self.g_rqname == '현재가조회':
                self.g_flag_6 = 1
            return

        if rqname == '증거금세부내역조회요청':
            temp = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
                                           trcode, '', rqname, 0,
                                           '100주문가능금액')  # 주문가능금액 저장
            self.g_ord_amt_possible = int(temp.strip())
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", screen_no)
            self.g_flag_1 = 1

        if rqname == '계좌평가현황요청':
            ii = 0

            user_id = None
            jongmok_cd = None
            jongmok_nm = None
            own_stock_cnt = 0
            buy_price = 0
            own_amt = 0

            repeat_cnt = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", [trcode, rqname])  # 보유종목수 가져오기

            self.write_msg_log("TB_ACCNT_INFO 테이블 설정 시작")
            self.write_msg_log("보유종목수 : " + str(repeat_cnt))

            while ii < repeat_cnt:
                user_id = ''
                jongmok_cd = ''
                own_stock_cnt = 0
                buy_price = 0
                own_amt = 0

                user_id = self.g_user_id
                jongmok_cd = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, '',
                                                     rqname, ii, '종목코드').strip()[1:7]
                jongmok_nm = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, '',
                                                     rqname, ii, '종목명').strip()
                own_stock_cnt = int(
                    self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, '', rqname,
                                            ii, '보유수량').strip())
                buy_price = int(
                    self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", [trcode, '', rqname,
                                                                                                     ii, '평균단가']).strip(
                        '0').rstrip('.'))
                own_amt = int(
                    self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, '', rqname,
                                            ii, '매입금액').strip())
                if own_stock_cnt == 0:  # 보유주식수가 0이라면 저장하지 않음
                    continue
                self.insert_tb_accnt_info(jongmok_cd, jongmok_nm, buy_price, own_stock_cnt, own_amt)  # 계좌정보 테이블에 저장

                ii += 1

            self.write_msg_log('TB_ACCNT_INFO 테이블 설정 완료')
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", screen_no)

            if len(prev_next) == 0:
                self.g_is_next = 0
            else:
                self.g_is_next = int(prev_next)

            self.g_flag_2 = 1

        if rqname == '호가조회':
            cnt = 0
            ii = 0
            l_buy_hoga = 0

            cnt = self.kiwoom.dynamicCall("GetRepeatCnt(QString, QString)", [trcode, rqname])

            while ii < cnt:
                l_buy_hoga = int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
                                                         [trcode, '', rqname, ii, '매수최우선호가']).strip())
                l_buy_hoga = abs(l_buy_hoga)
                ii += 1

            self.g_buy_hoga = l_buy_hoga
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", screen_no)
            self.g_flag_7 = 1

        if rqname == '현재가조회':
            self.g_cur_price = int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
                                                           [trcode, '', rqname, 0, '현재가']).strip())
            self.g_cur_price = abs(self.g_cur_price)
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", screen_no)
            self.g_flag_6 = 1

    # 주식주문 요청 이벤트 메소드
    def axKHOpenAPI1_OnReceiveMsg(self, screen_no, rqname, trcode, msg):
        if rqname == '매수주문':
            self.write_msg_log('===매수주문 원장 응답정보 출력 시작===')
            self.write_msg_log('sScrNo : [' + screen_no + ']')
            self.write_msg_log('sRQName : [' + rqname + ']')
            self.write_msg_log('sTrCode : [' + trcode + ']')
            self.write_msg_log('sMsg : [' + msg + ']')
            self.write_msg_log('===매수주문 원장 응답정보 출력 종료===')
            self.g_flag_3 = 1  # 매수주문 응답완료 설정

        if rqname == '매도주문':
            self.write_msg_log('===매도주문 원장 응답정보 출력 시작===')
            self.write_msg_log('sScrNo : [' + screen_no + ']')
            self.write_msg_log('sRQName : [' + rqname + ']')
            self.write_msg_log('sTrCode : [' + trcode + ']')
            self.write_msg_log('sMsg : [' + msg + ']')
            self.write_msg_log('===매도주문 원장 응답정보 출력 종료===')
            self.g_flag_4 = 1  # 매도주문 응답완료 설정

        if rqname == '매도취소주문':
            self.write_msg_log('===매도취소주문 원장 응답정보 출력 시작===')
            self.write_msg_log('sScrNo : [' + screen_no + ']')
            self.write_msg_log('sRQName : [' + rqname + ']')
            self.write_msg_log('sTrCode : [' + trcode + ']')
            self.write_msg_log('sMsg : [' + msg + ']')
            self.write_msg_log('===매도취소주문 원장 응답정보 출력 종료===')
            self.g_flag_5 = 1  # 매도취소주문 응답완료 설정

    # 주식주문 내역, 체결내역 이벤트 메소드
    def axKHOpenAPI1_OnReceiveChejanData(self, gubun, it_cnt, fid_list):
        if gubun == '0':  # 주문내역 및 체결내역 수신
            chejan_gb = self.kiwoom.dynamicCall("GetChejanData(int)", 913).strip()  # 주문내역인지 체결내역인지 가져옴
            if chejan_gb == '접수':  # '접수'--->주문내역
                user_id = self.g_user_id
                jongmok_cd = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).strip()[1:7]
                jongmok_nm = self.get_jongmok_nm(jongmok_cd)
                ord_gb = self.kiwoom.dynamicCall("GetChejanData(int)", 907).strip()
                ord_no = self.kiwoom.dynamicCall("GetChejanData(int)", 9203).strip()
                org_ord_no = self.kiwoom.dynamicCall("GetChejanData(int)", 904).strip()
                ord_price = int(self.kiwoom.dynamicCall("GetChejanData(int)", 901).strip())
                ord_stock_cnt = int(self.kiwoom.dynamicCall("GetChejanData(int)", 900).strip())
                ord_amt = ord_price * ord_stock_cnt

                now = QDate.currentDate()
                ref_dt = now.toString('yyyyMMdd')
                ord_dtm = ref_dt + self.kiwoom.dynamicCall("GetChejanData(int)", 908).strip()

                self.write_msg_log('종목코드 : [' + jongmok_cd + ']')
                self.write_msg_log('종목명 : [' + jongmok_nm + ']')
                self.write_msg_log('주문구분 : [' + ord_gb + ']')
                self.write_msg_log('주문번호 : [' + ord_no + ']')
                self.write_msg_log('원주문번호 : [' + org_ord_no + ']')
                self.write_msg_log('주문금액 : [' + str(ord_price) + ']')
                self.write_msg_log('주문주식수 : [' + str(ord_stock_cnt) + ']')
                self.write_msg_log('주문금액 : [' + str(ord_amt) + ']')
                self.write_msg_log('주문일시 : [' + ord_dtm + ']')

                self.insert_tb_ord_lst(ref_dt, jongmok_cd, jongmok_nm, ord_gb, ord_no, org_ord_no, ord_price,
                                       ord_stock_cnt, ord_amt, ord_dtm)  # 주문내역 저장

                if ord_gb == '2':  # 매수주문일 경우 주문가능금액 갱신
                    self.update_tb_accnt(ord_gb, ord_amt)
            elif chejan_gb == '체결':
                user_id = self.g_user_id
                jongmok_cd = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).strip()[1:7]
                jongmok_nm = self.get_jongmok_nm(jongmok_cd)
                chegyul_gb = self.kiwoom.dynamicCall("GetChejanData(int)", 907).strip()  # 2:매수 1:매도
                chegyul_no = int(self.kiwoom.dynamicCall("GetChejanData(int)", 909).strip())
                chegyul_price = int(self.kiwoom.dynamicCall("GetChejanData(int)", 910).strip())
                chegyul_cnt = int(self.kiwoom.dynamicCall("GetChejanData(int)", 911).strip())
                chegyul_amt = chegyul_price * chegyul_cnt
                org_ord_no = self.kiwoom.dynamicCall("GetChejanData(int)", 904).strip()

                now = QDate.currentDate()
                ref_dt = now.toString('yyyyMMdd')
                chegyul_dtm = ref_dt + self.kiwoom.dynamicCall("GetChejanData(int)", 908).strip()
                ord_no = self.kiwoom.dynamicCall("GetChejanData(int)", 9203).strip()

                self.write_msg_log('종목코드 : [' + jongmok_cd + ']')
                self.write_msg_log('종목명 : [' + jongmok_nm + ']')
                self.write_msg_log('체결구분 : [' + chegyul_gb + ']')
                self.write_msg_log('체결번호 : [' + str(chegyul_no) + ']')
                self.write_msg_log('체결가 : [' + str(chegyul_price) + ']')
                self.write_msg_log('체결주식수 : [' + str(chegyul_cnt) + ']')
                self.write_msg_log('체결금액 : [' + str(chegyul_amt) + ']')
                self.write_msg_log('체결일시 : [' + chegyul_dtm + ']')
                self.write_msg_log('주문번호 : [' + ord_no + ']')
                self.write_msg_log('원주문번호 : [' + org_ord_no + ']')

                self.insert_tb_chegyul_lst(ref_dt, jongmok_cd, jongmok_nm, chegyul_gb, chegyul_no, chegyul_price,
                                           chegyul_cnt, chegyul_amt, chegyul_dtm, ord_no, org_ord_no)  # 체결내역 저장

                if chegyul_gb == '1':  # 매도체결이라면 계좌 테이블의 매수가능금액을 늘려줌
                    self.update_tb_accnt(chegyul_gb, chegyul_amt)
        if gubun == '1':  # 계좌정보 수신
            user_id = self.g_user_id
            jongmok_cd = self.kiwoom.dynamicCall("GetChejanData(int)", 9001).strip()[1:7]
            boyu_cnt = int(self.kiwoom.dynamicCall("GetChejanData(int)", 930).strip())
            boyu_price = int(self.kiwoom.dynamicCall("GetChejanData(int)", 931).strip())
            boyu_amt = int(self.kiwoom.dynamicCall("GetChejanData(int)", 932).strip())

            l_jongmok_nm = self.get_jongmok_nm(jongmok_cd)

            self.write_msg_log('종목코드 : [' + jongmok_cd + ']')
            self.write_msg_log('보유주식수 : [' + str(boyu_cnt) + ']')
            self.write_msg_log('보유가 : [' + str(boyu_price) + ']')
            self.write_msg_log('보유금액 : [' + str(boyu_amt) + ']')

            self.merge_tb_accnt_info(jongmok_cd, l_jongmok_nm, boyu_cnt, boyu_price, boyu_amt)  # 계좌정보 저장

    # 매수가능금액 요청
    @pyqtSlot()
    def set_tb_accnt(self):
        for_cnt = 0
        for_flag = 0

        self.write_msg_log('TB_ACCNT 테이블 세팅 시작')
        self.g_ord_amt_possible = 0  # 매수가능금액

        for_flag = 0
        while True:
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.g_accnt_no)
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")

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

        try:
            sql_insert = "MERGE INTO TB_ACCNT a USING(SELECT NVL(MAX(USER_ID), ' ') USER_ID, NVL(MAX(ACCNT_NO), " \
                         "' ') ACCNT_NO, NVL(MAX(REF_DT), ' ') REF_DT FROM TB_ACCNT WHERE USER_ID = :1 AND ACCNT_NO = " \
                         ":2 AND REF_DT = :3) b ON (a.USER_ID = b.USER_ID AND a.ACCNT_NO = b.ACCNT_NO AND a.REF_DT = " \
                         "b.REF_DT) WHEN MATCHED THEN UPDATE SET ORD_POSSIBLE_AMT = :4, UPDT_DTM = :5, " \
                         "UPDT_ID = 'ats' WHEN NOT MATCHED THEN INSERT (a.USER_ID, a.ACCNT_NO, a.REF_DT, " \
                         "a.ORD_POSSIBLE_AMT, a.INST_DTM, a.INST_ID) VALUES(:6, :7, :8, :9, :10, 'ats') "

            cur.execute(sql_insert, (self.g_user_id,
                                     self.g_accnt_no, now.toString('yyyyMMdd'),
                                     self.g_ord_amt_possible, datetime.datetime.now(),
                                     self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd'),
                                     self.g_ord_amt_possible, datetime.datetime.now()))
            conn.commit()
        except Exception as ex:
            self.write_err_log("MERGE_TB_ACCNT ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 계좌정보 요청
    @pyqtSlot()
    def set_tb_accnt_info(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()
        for_cnt = 0
        for_flag = 0
        self.g_flag_2 = 1

        try:
            sql_insert = "DELETE FROM TB_ACCNT_INFO WHERE REF_DT =:1 AND USER_ID = :2"  # 당일 기준 계좌정보 삭제
            cur.execute(sql_insert, (now.toString('yyyyMMdd'), self.g_user_id))
            conn.commit()
        except Exception as ex:
            self.write_err_log("DELETE TB_ACCNT_INFO ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

        self.g_is_next = 0
        while True:
            # for_flag = 0
            while True:
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.g_accnt_no)
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "")
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "상장폐지조회구분", "1")
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")

                # self.g_flag_2 = 0 # ???????(이거 있으면 무한 루프)
                self.g_rqname = '계좌평가현황요청'

                scr_no = self.get_scr_no()

                # 계좌정보 데이터 수신 요청. 이벤트 발생
                self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "계좌평가현황요청", "OPW00004",
                                        self.g_is_next, scr_no)

                for_cnt = 0
                # 응답 대기 ------------------> 쓸모없는 코드? 응답 대기하면 이벤트 발생하지 않고 루프를 다시 돈다. 그 이후 정상작동
                while True:
                    if self.g_flag_2 == 1:
                        time.sleep(1)
                        self.kiwoom.dynamicCall("DisconnectRealData(QString)", scr_no)
                        for_flag = 1
                        break
                    else:
                        time.sleep(1)
                        for_cnt += 1
                        if for_cnt == 5:
                            for_flag = 0
                            break
                        else:
                            continue

                time.sleep(1)
                self.kiwoom.dynamicCall("DisconnectRealData(QString)", scr_no)

                if for_flag == 1:
                    break
                elif for_flag == 0:
                    time.sleep(1)
                    continue

            if self.g_is_next == 0:
                break
            time.sleep(1)

    # 계좌정보 테이블 설정
    def insert_tb_accnt_info(self, i_jongmok_cd, i_jongmok_nm, i_buy_price, i_own_stock_cnt, i_own_amt):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        try:
            sql_insert = "INSERT INTO TB_ACCNT_INFO VALUES(:1, :2, :3, :4, :5, :6, :7, :8, 'ats', :9, NULL, NULL)"
            cur.execute(sql_insert, (
                self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd'), i_jongmok_cd, i_jongmok_nm, i_buy_price,
                i_own_stock_cnt, i_own_amt, datetime.datetime.now()))
            conn.commit()
        except Exception as ex:
            self.write_err_log("INSERT TB_ACCNT_INFO ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 주문내역 각 항목의 값을 TB_ORD_LST 테이블에 저장
    def insert_tb_ord_lst(self, i_ref_dt, i_jongmok_cd, i_jongmok_nm, i_ord_gb, i_ord_no, i_org_ord_no, i_ord_price,
                          i_ord_stock_cnt, i_ord_amt, i_ord_dtm):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        try:
            sql_insert = "INSERT INTO TB_ORD_LST VALUES(:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, 'ats', :13, NULL, NULL)"
            cur.execute(sql_insert, (
                self.g_user_id, self.g_accnt_no, i_ref_dt, i_jongmok_cd, i_jongmok_nm, i_ord_gb, i_ord_no, i_org_ord_no,
                i_ord_price, i_ord_stock_cnt, i_ord_amt, i_ord_dtm, datetime.datetime.now()))
            conn.commit()
        except Exception as ex:
            self.write_err_log("INSERT TB_ORD_LST ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 주문이 완료되면 주문가능금액을 조정
    def update_tb_accnt(self, i_chegyul_gb, i_chegyul_amt):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        if i_chegyul_gb == '2':  # 매수일 때 주문가능금액에서 체결금액 빼기
            try:
                sql_insert = "UPDATE TB_ACCNT SET ORD_POSSIBLE_AMT = ORD_POSSIBLE_AMT - :1, UPDT_DTM = :2, UPDT_ID = 'ats' WHERE USER_ID = :3 AND ACCNT_NO = :4 AND REF_DT = :5"
                cur.execute(sql_insert, (
                    i_chegyul_amt, datetime.datetime.now(), self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd')))
                conn.commit()
            except Exception as ex:
                self.write_err_log("UPDATE TB_ACCNT ex.Message : [" + str(ex) + "]")
        elif i_chegyul_gb == '1':  # 매도일 때 주문가능금액에 체결금액 더하기
            try:
                sql_insert = "UPDATE TB_ACCNT SET ORD_POSSIBLE_AMT = ORD_POSSIBLE_AMT + :1, UPDT_DTM = :2, UPDT_ID = 'ats' WHERE USER_ID = :3 AND ACCNT_NO = :4 AND REF_DT = :5"
                cur.execute(sql_insert, (
                    i_chegyul_amt, datetime.datetime.now(), self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd')))
                conn.commit()
            except Exception as ex:
                self.write_err_log("UPDATE TB_ACCNT ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 체결내역 각 항목을 TB_CHEGYUL_LST 테이블에 저장
    def insert_tb_chegyul_lst(self, i_ref_dt, i_jongmok_cd, i_jongmok_nm, i_chegyul_gb, i_chegyul_no, i_chegyul_price,
                              i_chegyul_stock_cnt, i_chegyul_amt, i_chegyul_dtm, i_ord_no, i_org_ord_no):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        try:
            sql_insert = "INSERT INTO TB_CHEGYUL_LST VALUES(:1, :2, :3, :4, :5, :6, :7, :8, :9, :10, :11, :12, :13, 'ats', :14, NULL, NULL)"
            cur.execute(sql_insert, (
                self.g_user_id, self.g_accnt_no, i_ref_dt, i_jongmok_cd, i_jongmok_nm, i_chegyul_gb, i_ord_no,
                i_chegyul_gb,
                i_chegyul_no, i_chegyul_price, i_chegyul_stock_cnt, i_chegyul_amt, i_chegyul_dtm,
                datetime.datetime.now()))
            conn.commit()
        except Exception as ex:
            self.write_err_log("INSERT TB_CHEGYUL_LST ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # 수신받은 계좌정보를 TB_ACCNT_INFO 테이블에 갱신 또는 삽입
    def merge_tb_accnt_info(self, i_jongmok_cd, i_jongmok_nm, i_boyu_cnt, i_boyu_price, i_boyu_amt):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        # 계좌정보 테이블 세팅, 기존에 보유한 종목이면 갱신, 보유하지 않았으면 신규로 삽입
        try:
            sql_insert = "MERGE INTO TB_ACCNT_INFO a USING(SELECT NVL(MAX(USER_ID), '0') USER_ID, NVL(MAX(REF_DT), " \
                         "'0') REF_DT, NVL(MAX(JONGMOK_CD), '0') JONGMOK_CD, NVL(MAX(JONGMOK_NM), '0') JONGMOK_NM " \
                         "FROM TB_ACCNT_INFO WHERE USER_ID = :1 AND ACCNT_NO = :2 AND JONGMOK_CD = :3 AND REF_DT = " \
                         ":4) b ON (a.USER_ID = b.USER_ID AND a.JONGMOK_CD = b.JONGMOK_CD AND a.REF_DT = b.REF_DT) " \
                         "WHEN MATCHED THEN UPDATE SET OWN_STOCK_CNT = :5, BUY_PRICE = :6, OWN_AMT = :7, UPDT_DTM = " \
                         ":8, UPDT_ID = :9 WHEN NOT MATCHED THEN INSERT(a.USER_ID, a.ACCNT_NO, a.REF_DT, " \
                         "a.JONGMOK_CD, a.JONGMOK_NM, a.BUY_PRICE, a.OWN_STOCK_CNT, a.OWN_AMT, a.INST_DTM, " \
                         "a.INST_ID) VALUES(:10, :11, :12, :13, :14, :15, :16, :17, :18, :19) "
            cur.execute(sql_insert, (
                self.g_user_id, self.g_accnt_no, i_jongmok_cd, now.toString('yyyyMMdd'), i_boyu_cnt, i_boyu_price,
                i_boyu_amt, datetime.datetime.now(), 'ats', self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd'),
                i_jongmok_cd, i_jongmok_nm, i_boyu_price, i_boyu_cnt, i_boyu_amt, datetime.datetime.now(), 'ats'))
            conn.commit()
        except Exception as ex:
            self.write_err_log("MERGE TB_ACCNT_INFO ex.Message : [" + str(ex) + "]")

        cur.close()
        conn.close()

    # TB_TRD_JONGMOK 테이블과 TB_ACCNT_INFO 테이블을 조인하여 매도대상 종목을 조회
    def sell_ord_first(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        # 두 테이블을 조인하여 매도대상 종목 조회
        sql_insert = "SELECT A.JONGMOK_CD, A.BUY_PRICE, A.OWN_STOCK_CNT, B.TARGET_PRICE FROM TB_ACCNT_INFO A, " \
                     "TB_TRD_JONGMOK B WHERE A.USER_ID = :1 AND A.ACCNT_NO = :2 AND A.REF_DT = :3 AND A.USER_ID = " \
                     "B.USER_ID AND A.JONGMOK_CD = B.JONGMOK_CD AND B.SELL_TRD_YN = 'Y' AND A.OWN_STOCK_CNT > 0 "
        cur.execute(sql_insert, (
            self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd')))

        for row in cur:
            l_jongmok_cd = str(row[0]).strip()
            l_buy_price = int(str(row[1]).strip())
            l_own_stock_cnt = int(str(row[2]).strip())
            l_target_price = int(str(row[3]).strip())

            self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
            self.write_msg_log('매입가 : [' + str(l_buy_price) + ']')
            self.write_msg_log('보유주식수 : [' + str(l_own_stock_cnt) + ']')
            self.write_msg_log('목표가 : [' + str(l_target_price) + ']')

            l_new_target_price = 0
            l_new_target_price = self.get_hoga_unit_price(l_target_price, l_jongmok_cd, 0)

            self.g_flag_4 = 0
            self.g_rqname = '매도주문'

            l_scr_no = self.get_scr_no()
            ret = 0

            # 매도주문 요청
            ret = self.kiwoom.dynamicCall(
                "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)", ['매도주문', l_scr_no,
                                                                                                   self.g_accnt_no, 2,
                                                                                                   l_jongmok_cd,
                                                                                                   l_own_stock_cnt,
                                                                                                   l_new_target_price,
                                                                                                   "00", ""])
            if ret == 0:
                self.write_msg_log('매도주문 Sendord() 호출 성공')
                self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
            else:
                self.write_msg_log('매도주문 Sendord() 호출 실패')
                self.write_msg_log('i_jongmok_cd : [' + l_jongmok_cd + ']')

            time.sleep(0.2)

            while True:
                if self.g_flag_4 == 1:
                    time.sleep(0.2)
                    self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
                    break
                else:
                    self.write_msg_log('매도주문 완료 대기 중')
                    time.sleep(0.2)
                    break
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)

        conn.commit()
        cur.close()
        conn.close()

    # 매도하기 전 유효한 호가가격단위 구하기
    def get_hoga_unit_price(self, i_price, i_jongmok_cd, i_hoga_unit_jump):
        l_rest = None
        l_market_type = 0

        try:
            l_market_type = int(
                str(self.kiwoom.dynamicCall("GetMarketType(QString)", i_jongmok_cd)))  # 시장구분 가져오기. 0:코스피, 10:코스닥
        except Exception as ex:
            self.write_err_log("get_hoga_unit_price() ex.Message : [" + str(ex) + "]")
        if i_price < 1000:
            return i_price + (i_hoga_unit_jump * 1)
        elif 1000 <= i_price < 5000:
            l_rest = i_price % 5
            if l_rest == 0:
                return i_price + (i_hoga_unit_jump * 5)
            elif l_rest < 3:
                return (i_price - l_rest) + (i_hoga_unit_jump * 5)
            else:
                return (i_price + (5 - l_rest)) + (i_hoga_unit_jump * 5)
        elif 5000 <= i_price < 10000:
            l_rest = i_price % 10
            if l_rest == 0:
                return i_price + (i_hoga_unit_jump * 10)
            elif l_rest < 5:
                return (i_price - l_rest) + (i_hoga_unit_jump * 10)
            else:
                return (i_price + (10 - l_rest)) + (i_hoga_unit_jump * 10)
        elif 10000 <= i_price < 50000:
            l_rest = i_price % 50
            if l_rest == 0:
                return i_price + (i_hoga_unit_jump * 50)
            elif l_rest < 25:
                return (i_price - l_rest) + (i_hoga_unit_jump * 50)
            else:
                return (i_price + (50 - l_rest)) + (i_hoga_unit_jump * 50)
        elif 50000 <= i_price < 100000:
            l_rest = i_price % 100
            if l_rest == 0:
                return i_price + (i_hoga_unit_jump * 100)
            elif l_rest < 50:
                return (i_price - l_rest) + (i_hoga_unit_jump * 100)
            else:
                return (i_price + (100 - l_rest)) + (i_hoga_unit_jump * 100)
        elif 100000 <= i_price < 500000:
            if l_market_type == 10:
                l_rest = i_price % 100
                if l_rest == 0:
                    return i_price + (i_hoga_unit_jump * 100)
                elif l_rest < 50:
                    return (i_price - l_rest) + (i_hoga_unit_jump * 100)
                else:
                    return (i_price + (100 - l_rest)) + (i_hoga_unit_jump * 100)
            else:
                l_rest = i_price % 500
                if l_rest == 0:
                    return i_price + (i_hoga_unit_jump * 500)
                elif l_rest < 250:
                    return (i_price - l_rest) + (i_hoga_unit_jump * 500)
                else:
                    return (i_price + (500 - l_rest)) + (i_hoga_unit_jump * 500)
        elif 500000 <= i_price:
            if l_market_type == 10:
                l_rest = i_price % 100
                if l_rest == 0:
                    return i_price + (i_hoga_unit_jump * 100)
                elif l_rest < 50:
                    return (i_price - l_rest) + (i_hoga_unit_jump * 100)
                else:
                    return (i_price + (100 - l_rest)) + (i_hoga_unit_jump * 100)
            else:
                l_rest = i_price % 1000
                if l_rest == 0:
                    return i_price + (i_hoga_unit_jump * 1000)
                elif l_rest < 500:
                    return (i_price - l_rest) + (i_hoga_unit_jump * 1000)
                else:
                    return (i_price + (1000 - l_rest)) + (i_hoga_unit_jump * 1000)
        return 0

    # 실시간 매수주문
    def real_buy_ord(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        # 두 테이블을 조인하여 매도대상 종목 조회
        sql_insert = "SELECT A.JONGMOK_CD, A.BUY_AMT, A.BUY_PRICE FROM TB_TRD_JONGMOK A WHERE A.USER_ID = :USER_ID AND A.BUY_TRD_YN = 'Y' ORDER BY A.PRIORITY"
        cur.execute(sql_insert, USER_ID=self.g_user_id)
        for row in cur:
            l_jongmok_cd = str(row[0]).strip()
            l_buy_amt = int(str(row[1]).strip())  # 매수금액
            l_buy_price = int(str(row[2]).strip())  # 매수가

            l_buy_price_tmp = self.get_hoga_unit_price(l_buy_price, l_jongmok_cd, 1)  # 매수호가 구하기
            l_buy_ord_stock_cnt = int(l_buy_amt / l_buy_price_tmp)  # 매수주문 주식

            self.write_msg_log('종목코드 : [' + str(l_jongmok_cd) + ']')
            self.write_msg_log('종목명 : [' + self.get_jongmok_nm(l_jongmok_cd) + ']')
            self.write_msg_log('매수금액 : [' + str(l_buy_amt) + ']')
            self.write_msg_log('매수가 : [' + str(l_buy_price_tmp) + ']')
            l_own_stock_cnt = 0
            l_own_stock_cnt = self.get_own_stock_cnt(l_jongmok_cd)  # 해당 종목 보유주식수 구하기
            self.write_msg_log('보유주식수 : [' + str(l_own_stock_cnt) + ']')
            # 이 부분 수정해서 태이블에서 수정하면 추가매수 할 수 있게 하기
            if l_own_stock_cnt > 0:
                self.write_msg_log('해당 종목을 보유 중이므로 매수하지 않음')
                continue

            l_buy_not_chegyul_yn = self.get_buy_not_chegyul_yn(l_jongmok_cd)  # 미체결 매수주문 여부 확인

            if l_buy_not_chegyul_yn == 'Y':  # 미체결 매수주문이 있으므로 매수하지 않음
                self.write_msg_log('해당 종목에 미체결 매수주문이 있으므로 매수하지 않음')
                continue

            l_for_flag = 0
            l_for_cnt = 0
            self.g_buy_hoga = 0
            while True:
                self.g_rqname = ''
                self.g_rqname = '호가조회'
                self.g_flag_7 = 1
                self.kiwoom.dynamicCall(
                    "SetInputValue(QString, QString)", ['종목코드', l_jongmok_cd])

                l_scr_no_2 = self.get_scr_no()
                self.kiwoom.dynamicCall(
                    "CommRqData(QString, QString, int, QString)", ['호가조회', 'opt10004', 0, l_scr_no_2])

                try:
                    l_for_cnt = 0
                    while True:
                        if self.g_flag_7 == 1:
                            time.sleep(0.2)
                            self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no_2)
                            l_for_flag = 1
                            break
                        else:
                            self.write_msg_log('호가조회 완료 대기 중')
                            time.sleep(0.2)
                            l_for_cnt += 1
                            if l_for_cnt == 5:
                                l_for_flag = 0
                                break
                            else:
                                continue
                except Exception as ex:
                    self.write_err_log("real_buy_ord() 호가조회 ex.Message : [" + str(ex) + "]")

                self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no_2)

                if l_for_flag == 1:
                    break
                elif l_for_flag == 0:
                    time.sleep(0.2)
                    continue
                time.sleep(0.2)

            if l_buy_price > self.g_buy_hoga:
                self.write_msg_log('해당 종목의 매수가가 최우선 매수호가보다 크므로 매수주문하지 않음')
                continue

            self.g_flag_3 = 0
            self.g_rqname = '매수주문'

            l_scr_no = self.get_scr_no()
            ret = 0

            # 매수주문 요청
            ret = self.kiwoom.dynamicCall(
                "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                ['매수주문', l_scr_no, self.g_accnt_no, 1, l_jongmok_cd, l_buy_ord_stock_cnt, l_buy_price, '00', ''])
            if ret == 0:
                self.write_msg_log('매수주문 Sendord() 호출 성공')
                self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
            else:
                self.write_msg_log('매수주문 Sendord() 호출 실패')
                self.write_msg_log('i_jongmok_cd : [' + l_jongmok_cd + ']')

            time.sleep(0.2)

            while True:
                if self.g_flag_3 == 1:
                    time.sleep(0.2)
                    self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
                    break
                else:
                    self.write_msg_log('매수주문 완료 대기 중')
                    time.sleep(0.2)
                    break
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
            time.sleep(1)

        conn.commit()
        cur.close()
        conn.close()

    # 보유주식수 구하기(이미 보유한 주식이라면 매수주문을 내지 않는다)
    def get_own_stock_cnt(self, i_jongmok_cd):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()
        l_own_stock_cnt = 0
        # 두 테이블을 조인하여 매도대상 종목 조회
        sql_insert = "SELECT NVL(MAX(OWN_STOCK_CNT), 0) OWN_STOCK_CNT FROM TB_ACCNT_INFO WHERE USER_ID = :1 AND JONGMOK_CD = :2 AND ACCNT_NO = :3 AND REF_DT = :4"
        cur.execute(sql_insert, (self.g_user_id, i_jongmok_cd, self.g_accnt_no, now.toString('yyyyMMdd')))

        for row in cur:
            l_own_stock_cnt = int(str(row[0]))  # 보유주식수 구하기

        conn.commit()
        cur.close()
        conn.close()

        return l_own_stock_cnt

    # 미체결 매수주문 여부 가져오기
    def get_buy_not_chegyul_yn(self, i_jongmok_cd):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()
        l_buy_not_chegyul_ord_stock_cnt = 0
        l_buy_not_chegyul_yn = None

        # 주문내역과 체결내역 테이블 조회
        sql_insert = "SELECT NVL(SUM(ORD_STOCK_CNT - CHEGYUL_STOCK_CNT), 0) BUY_NOT_CHEGYUL_ORD_STOCK_CNT FROM(SELECT ORD_STOCK_CNT ORD_STOCK_CNT, (SELECT NVL(MAX(b.CHEGYUL_STOCK_CNT), 0) CHEGYUL_STOCK_CNT FROM TB_CHEGYUL_LST b WHERE b.USER_ID = a.USER_ID AND b.ACCNT_NO = a.ACCNT_NO AND b.REF_DT = a.REF_DT AND b.JONGMOK_CD = a.JONGMOK_CD AND b.ORD_GB = a.ORD_GB AND b.ORD_NO = a.ORD_NO) CHEGYUL_STOCK_CNT FROM TB_ORD_LST a WHERE a.REF_DT = :1 AND a.USER_ID = :2 AND a.ACCNT_NO = :3 AND a.JONGMOK_CD = :4 AND a.ORD_GB = :5 AND a.ORG_ORD_NO = :6 AND NOT EXISTS(SELECT '1' FROM TB_ORD_LST b WHERE b.USER_ID = a.USER_ID AND b.ACCNT_NO = a.ACCNT_NO AND b.REF_DT = a.REF_DT AND b.JONGMOK_CD = a.JONGMOK_CD AND b.ORD_GB = a.ORD_GB AND b.ORG_ORD_NO = a.ORD_NO))x"
        cur.execute(sql_insert,
                    (now.toString('yyyyMMdd'), self.g_user_id, self.g_accnt_no, i_jongmok_cd, '2', '0000000'))

        for row in cur:
            l_buy_not_chegyul_ord_stock_cnt = int(str(row[0]))  # 미체결 매수주문 주식수 구하기

        conn.commit()
        cur.close()
        conn.close()

        if l_buy_not_chegyul_ord_stock_cnt > 0:
            l_buy_not_chegyul_yn = 'Y'
        else:
            l_buy_not_chegyul_yn = 'N'

        return l_buy_not_chegyul_yn

    # 실시간 매도대상 종목 조회
    def real_sell_ord(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        l_target_price = 0
        l_own_stock_cnt = 0
        self.write_msg_log('real_sell_ord 시작')

        # 거래정보 및 계좌정보 테이블 조회
        sql_insert = "SELECT A.JONGMOK_CD, A.TARGET_PRICE, B.OWN_STOCK_CNT FROM TB_TRD_JONGMOK A, TB_ACCNT_INFO B " \
                     "WHERE A.USER_ID = :1 AND A.JONGMOK_CD = B.JONGMOK_CD AND B.ACCNT_NO = :2 AND B.REF_DT = :3 " \
                     "AND A.SELL_TRD_YN = :4 AND B.OWN_STOCK_CNT > :5 "
        cur.execute(sql_insert, (self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd'), 'Y', 0))

        for row in cur:
            l_jongmok_cd = str(row[0]).strip()
            l_target_price = int(str(row[1]).strip())
            l_own_stock_cnt = int(str(row[2]).strip())

            self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
            self.write_msg_log('종목명 : [' + self.get_jongmok_nm(l_jongmok_cd) + ']')
            self.write_msg_log('목표가 : [' + str(l_target_price) + ']')
            self.write_msg_log('보유주식수 : [' + str(l_own_stock_cnt) + ']')

            l_sell_not_chegyul_ord_stock_cnt = self.get_sell_not_chegyul_ord_stock_cnt(l_jongmok_cd)  # 미체결 매도주문 주식수 구하기

            if l_sell_not_chegyul_ord_stock_cnt == l_own_stock_cnt:  # 미체결 매도주문 주식수와 보유주식수가 같으면 기 주문종목이므로 매도주문하지 않음
                continue
            else:  # 미체결 매도주문 주식수와 보유주식수가 같지 않으면 아직 매도하지 않은 종목임
                l_sell_ord_stock_cnt_tmp = l_own_stock_cnt - l_sell_not_chegyul_ord_stock_cnt  # 보유주식수에서 미체결 매도주문 주식수를 빼서 매도주문 주식수를 구함
                if l_sell_ord_stock_cnt_tmp <= 0:  # 매도대상 주식수가 0 이하라면 매도하지 않음
                    continue

                l_new_target_price = self.get_hoga_unit_price(l_target_price, l_jongmok_cd, 0)  # 매도호기를 구함

                self.g_flag_4 = 0
                self.g_rqname = '매도주문'

                l_scr_no = self.get_scr_no()

                # 매도주문 요청
                ret = self.kiwoom.dynamicCall(
                    "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                    ['매도주문', l_scr_no, self.g_accnt_no, 2, l_jongmok_cd, l_sell_ord_stock_cnt_tmp, l_new_target_price,
                     '00', ''])

                if ret == 0:
                    self.write_msg_log('매도주문 Sendord() 호출 성공')
                    self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
                else:
                    self.write_msg_log('매도주문 Sendord() 호출 실패')
                    self.write_msg_log('i_jongmok_cd : [' + l_jongmok_cd + ']')
                time.sleep(0.2)

                while True:
                    if self.g_flag_4 == 1:
                        time.sleep(0.2)
                        self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
                        break
                    else:
                        self.write_msg_log('매도주문 완료 대기중')
                        time.sleep(0.2)
                        break
                self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)

        conn.commit()
        cur.close()
        conn.close()

    # 이미 매도주문한 종목인지 확인하기 위해 미체결 매도주문을 확인
    def get_sell_not_chegyul_ord_stock_cnt(self, i_jongmok_cd):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        l_sell_not_chegyul_ord_stock_cnt = 0

        # 주문내역과 체결내역 테이블 조회
        sql_insert = "SELECT NVL(SUM(ORD_STOCK_CNT - CHEGYUL_STOCK_CNT), 0) SELL_NOT_CHEGYUL_ORD_STOCK_CNT FROM (" \
                     "SELECT ORD_STOCK_CNT ORD_STOCK_CNT, (SELECT NVL(MAX(b.CHEGYUL_STOCK_CNT), 0) CHEGYUL_STOCK_CNT " \
                     "FROM TB_CHEGYUL_LST B WHERE b.USER_ID = a.USER_ID AND b.ACCNT_NO = a.ACCNT_NO AND b.REF_DT = " \
                     "a.REF_DT AND b.JONGMOK_CD = a.JONGMOK_CD AND b.ORD_GB = a.ORD_GB AND b.ORD_NO = a.ORD_NO) " \
                     "CHEGYUL_STOCK_CNT FROM TB_ORD_LST a WHERE a.REF_DT = :1 AND a.USER_ID = :2 AND a.JONGMOK_CD = " \
                     ":3 AND a.ACCNT_NO = :4 AND a.ORD_GB = '1' AND a.ORG_ORD_NO = '0000000' AND NOT EXISTS (SELECT " \
                     "'1' FROM TB_ORD_LST b WHERE b.USER_ID = a.USER_ID AND b.ACCNT_NO = a.ACCNT_NO AND b.REF_DT = " \
                     "a.REF_DT AND b.JONGMOK_CD = a.JONGMOK_CD AND b.ORD_GB = a.ORD_GB AND b.ORG_ORD_NO = a.ORD_NO)) "
        cur.execute(sql_insert, (now.toString('yyyyMMdd'), self.g_user_id, i_jongmok_cd, self.g_accnt_no))

        for row in cur:
            l_sell_not_chegyul_ord_stock_cnt = int(str(row[0]))  # 미체결 매도주문 주식수 가져오기

        conn.commit()
        cur.close()
        conn.close()

        return l_sell_not_chegyul_ord_stock_cnt

    # 실시간 손절주문
    def real_cut_loss_ord(self):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        # 거래정보 및 계좌정보 테이블 조회
        sql_insert = "SELECT A.JONGMOK_CD, A.CUT_LOSS_PRICE, B.OWN_STOCK_CNT FROM TB_TRD_JONGMOK A, TB_ACCNT_INFO B " \
                     "WHERE A.USER_ID = :1 AND A.JONGMOK_CD = B.JONGMOK_CD AND B.ACCNT_NO = :2 AND B.REF_DT = :3 " \
                     "AND A.SELL_TRD_YN = :4 AND B.OWN_STOCK_CNT > :5 "
        cur.execute(sql_insert, (self.g_user_id, self.g_accnt_no, now.toString('yyyyMMdd'), 'Y', 0))

        for row in cur:
            l_jongmok_cd = str(row[0]).strip()
            l_cut_loss_price = int(str(row[1]).strip())
            l_own_stock_cnt = int(str(row[2]).strip())

            self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
            self.write_msg_log('종목명 : [' + self.get_jongmok_nm(l_jongmok_cd) + ']')
            self.write_msg_log('손절가 : [' + str(l_cut_loss_price) + ']')
            self.write_msg_log('보유주식수 : [' + str(l_own_stock_cnt) + ']')

            l_for_flag = 0
            self.g_cur_price = 0

            while True:
                self.g_rqname = '현재가조회'
                self.g_flag_6 = 1
                self.kiwoom.dynamicCall("SetInputValue(QString, QString)", ['종목코드', l_jongmok_cd])

                l_scr_no = self.get_scr_no()

                # 현재가 조회 요청
                self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)",
                                        [self.g_rqname, 'opt10001', 0, l_scr_no])
                try:
                    l_for_cnt = 0
                    while True:
                        if self.g_flag_6 == 1:
                            time.sleep(0.2)
                            self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
                            l_for_flag = 1
                            break
                        else:
                            self.write_msg_log('현재가조회 완료 대기 중')
                            time.sleep(0.2)
                            l_for_cnt += 1
                            if l_for_cnt == 5:
                                l_for_flag = 0
                                break
                            else:
                                continue
                except Exception as ex:
                    self.write_err_log("real_cut_loss_ord() 현재가조회 ex.Message : [" + str(ex) + "]")

                self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)

                if l_for_flag == 1:
                    break
                elif l_for_flag == 0:
                    time.sleep(0.2)
                    continue
                time.sleep(0.2)

            if self.g_cur_price < l_cut_loss_price:  # 현재가가 손절가 이탈 시
                self.sell_canc_ord(l_jongmok_cd)

                self.g_flag_4 = 0
                self.g_rqname = '매도주문'

                l_scr_no = self.get_scr_no()

                # 매도주문 요청
                ret = self.kiwoom.dynamicCall(
                    "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                    ['매도주문', l_scr_no, self.g_accnt_no, 2, l_jongmok_cd, l_own_stock_cnt, 0,
                     '03', ''])
                if ret == 0:
                    self.write_msg_log('매도주문 Sendord() 호출 성공')
                    self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
                else:
                    self.write_msg_log('매도주문 Sendord() 호출 실패')
                    self.write_msg_log('i_jongmok_cd : [' + l_jongmok_cd + ']')
                time.sleep(0.2)

                while True:
                    if self.g_flag_4 == 1:
                        time.sleep(0.2)
                        self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
                        break
                    else:
                        self.write_msg_log('매도주문 완료 대기중')
                        time.sleep(0.2)
                        break
                self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)

                self.update_tb_trd_jongmok(l_jongmok_cd)

        conn.commit()
        cur.close()
        conn.close()

    # 현재가가 손절가를 이탈하면 손절주문
    def sell_canc_ord(self, i_jongmok_cd):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        sql_insert = "SELECT ROWID RID, JONGMOK_CD, (ORD_STOCK_CNT - (SELECT NVL(MAX(b.CHEGYUL_STOCK_CNT), " \
                     "0) CHEGYUL_STOCK_CNT FROM TB_CHEGYUL_LST b WHERE b.USER_ID = a.USER_ID AND b.ACCNT_NO = " \
                     "a.ACCNT_NO AND b.REF_DT = a.REF_DT AND b.JONGMOK_CD = a.JONGMOK_CD AND b.ORD_GB = a.ORD_GB AND " \
                     "b.ORD_NO = a.ORD_NO)) SELL_NOT_CHEGYUL_ORD_STOCK_CNT, ORD_PRICE, ORD_NO, ORG_ORD_NO FROM " \
                     "TB_ORD_LST a WHERE a.REF_DT = :1 AND a.USER_ID = :2 AND a.ACCNT_NO = :3 AND a.JONGMOK_CD = :4 " \
                     "AND " \
                     "a.ORD_GB = :5 AND a.ORG_ORD_NO = :6 AND NOT EXISTS (SELECT '1' FROM TB_ORD_LST b WHERE " \
                     "b.USER_ID = a.USER_ID AND b.ACCNT_NO = a.ACCNT_NO AND b.REF_DT = a.REF_DT AND b.JONGMOK_CD = " \
                     "a.JONGMOK_CD AND b.ORD_GB = a.ORD_GB AND b.ORG_ORD_NO = a.ORD_NO) "
        cur.execute(sql_insert,
                    (now.toString('yyyyMMdd'), self.g_user_id, self.g_accnt_no, i_jongmok_cd, '1', '0000000'))


        for row in cur:
            l_rid = str(row[0]).strip()
            l_jongmok_cd = str(row[1]).strip()
            l_ord_stock_cnt = int(str(row[2]).strip())
            l_ord_price = int(str(row[3]).strip())
            l_ord_no = str(row[4]).strip()
            l_org_ord_no = str(row[5]).strip()

            self.g_flag_5 = 0
            self.g_rqname = '매도취소주문'

            l_scr_no = self.get_scr_no()

            # 매도취소주문 요청
            ret = self.kiwoom.dynamicCall(
                "SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                ['매도취소주문', l_scr_no, self.g_accnt_no, 4, l_jongmok_cd, l_ord_stock_cnt, 0,
                 '03', l_ord_no])

            if ret == 0:
                self.write_msg_log('매도취소주문 Sendord() 호출 성공')
                self.write_msg_log('종목코드 : [' + l_jongmok_cd + ']')
            else:
                self.write_msg_log('매도취소주문 Sendord() 호출 실패')
                self.write_msg_log('i_jongmok_cd : [' + l_jongmok_cd + ']')
            time.sleep(0.2)

            while True:
                if self.g_flag_5 == 1:
                    time.sleep(0.2)
                    self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
                    break
                else:
                    self.write_msg_log('매도취소주문 완료 대기중')
                    time.sleep(0.2)
                    break
            self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)

            time.sleep(1)

        conn.commit()
        cur.close()
        conn.close()

    # 손절주문시 매수거래여부를 'N'으로 설정정
    def update_tb_trd_jongmok(self, i_jongmok_cd):
        conn = cx_Oracle.connect('ats', '1234', 'localhost:1521/xe', encoding='UTF-8', nencoding='UTF-8')
        cur = conn.cursor()
        now = QDate.currentDate()

        try:
            sql_insert = "UPDATE TB_TRD_JONGMOK SET BUY_TRD_YN = 'N', UPDT_DTM = :1, UPDT_ID = 'ats' WHERE USER_ID = :2 AND JONGMOK_CD = :3"
            cur.execute(sql_insert,
                        (datetime.datetime.now(), self.g_user_id, i_jongmok_cd))
        except Exception as ex:
            self.write_err_log("UPDATE TB_TRD_JONGMOK ex.Message : [" + str(ex) + "]")

        conn.commit()
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
