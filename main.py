import sys
import cx_Oracle

from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5 import uic
from PyQt5.QtGui import *
from PyQt5.QAxContainer import *

# UI파일 연결
# 단, UI파일은 Python 코드 파일과 같은 디렉토리에 위치해야한다.
form_class = uic.loadUiType("untitled.ui")[0]


# 화면을 띄우는데 사용되는 Class 선언
class WindowClass(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.g_scr_no = 0  # Open API 요청번호

        # 키움증권 클래스를 사용하기 위해 인스턴스 생성(ProgID를 사용)
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        # 로그인창 띄우기, OnEventConnect 이벤트 발생시킴
        self.kiwoom.dynamicCall("CommConnect()")

        self.text_edit = QTextEdit(self)
        self.text_edit.setGeometry(10, 60, 280, 80)
        self.text_edit.setEnabled(False)

        # 이벤트(oneventconnect)와 이벤트 처리 메소드 연결
        self.kiwoom.OnEventConnect.connect(self.event_connect)

        self.statusBar.showMessage('Ready')

    # 이벤트 처리 함수
    def event_connect(self, err_code):
        if err_code == 0:
            self.text_edit.append("로그인 성공")
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
        db = cx_Oracle.connect("hr", "1234", dsn)

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


if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()
    myWindow.write_msg_log("asdf")

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
