import sys
import os
import cx_Oracle
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
LOCATION = r"C:\instantclient_19_9" # 32bit 버전 써야 한다
os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]

# 오라클 한글 지원
os.putenv('NLS_LANG', '.UTF8')


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
        login_status = QAction('아이디 / 계좌정보 가져오기', self)
        login_status.triggered.connect(self.menu_login_status)
        self.menu_login.addAction(login_act)
        self.menu_login.addAction(login_status)

        logout_act = QAction('로그아웃', self)
        logout_act.triggered.connect(self.menu_logout_act)  # 로그아웃 메소드 호출
        self.menu_logout.addAction(logout_act)

        # 콤보박스에서 선택한 계좌번호를 현재 계좌번호로 설정
        self.comboBox.currentIndexChanged.connect(self.comboBoxFunction)

        # 조회 버튼
        self.pushButton.clicked.connect(self.pushbutton_clicked)

    # 이벤트 처리 함수
    def event_connect(self, err_code):
        if err_code == 0:
            self.write_msg_log("로그인 성공")

    # 현재시각 가져오기
    def get_cur_tm(self):
        time = QTime.currentTime()
        answer = time.toString('hh.mm.ss')

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
        cul_num = 0
        self.tableWidget.clear()

        for row in rows:
            jongmok_cd = row['JONGMOK_CD']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(jongmok_cd))
            cul_num += 1
            jongmok_nm = row['JONGMOK_NM']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(jongmok_nm))
            cul_num += 1
            priority = row['PRIORITY']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(priority))
            cul_num += 1
            buy_amt = row['BUY_AMT']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(buy_amt))
            cul_num += 1
            buy_price = row['BUY_PRICE']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(buy_price))
            cul_num += 1
            target_price = row['TARGET_PRICE']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(target_price))
            cul_num += 1
            cut_loss_price = row['CUT_LOSS_PRICE']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(cut_loss_price))
            cul_num += 1
            buy_trd_yn = row['BUY_TRD_YN']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(buy_trd_yn))
            cul_num += 1
            sell_trd_yn = row['SELL_TRD_YN']
            self.tableWidget.setItem(row_num, cul_num, QTableWidgetItem(sell_trd_yn))
            cul_num = 0
            row_num += 1

        self.write_msg_log('TB_TRD_JONGMOK 테이블이 조회되었습니다')


if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
