import sys
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



if __name__ == "__main__":
    # QApplication : 프로그램을 실행시켜주는 클래스
    app = QApplication(sys.argv)

    # WindowClass의 인스턴스 생성
    myWindow = WindowClass()

    # 프로그램 화면을 보여주는 코드
    myWindow.show()

    # 프로그램을 이벤트루프로 진입시키는(프로그램을 작동시키는) 코드
    app.exec_()
