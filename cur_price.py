import sys
import time
from PyQt5.QAxContainer import *


class Cur_price():
    def __init__(self):
        self.g_rqname = ''
        self.g_flag_6 = 0
        self.g_scr_no = 0  # Open API 요청번호
        self.g_cur_price2 = sys.maxsize

        # 키움증권 클래스를 사용하기 위해 인스턴스 생성(ProgID를 사용)
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")

        # 키움증권 open api 응답 대기 이벤트
        # self.kiwoom.OnReceiveTrData.connect(self.axKHOpenAPI1_OnReceiveTrData)
        # self.kiwoom.OnReceiveRealData.connect(self.axKHOpenAPI1_OnReceiveRealData)

    # def axKHOpenAPI1_OnReceiveTrData(self, screen_no, rqname, trcode, recordname, prev_next):
    #     if self.g_rqname == rqname:  # 요청한 요청명과 Open API로부터 응답받은 요청명이 같다면
    #         pass
    #     else:
    #         if self.g_rqname == '현재가조회':
    #             self.g_flag_6 = 1
    #         return
    #
    #     if rqname == '현재가조회':
    #         self.g_cur_price = int(self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)",
    #                                                        [trcode, '', rqname, 0, '현재가']).strip())
    #         self.g_cur_price = abs(self.g_cur_price)
    #         self.kiwoom.dynamicCall("DisconnectRealData(QString)", screen_no)
    #         self.g_flag_6 = 1

    # def axKHOpenAPI1_OnReceiveRealData(self, jongmok_cd, real_ty, real_dt):
    #     self.g_cur_price = int(self.kiwoom.dynamicCall("GetCommRealData(QString, int)",
    #                                                    [jongmok_cd, 10]).strip())
    #     self.g_cur_price = abs(self.g_cur_price)
    #     self.g_flag_6 = 1

    # 요청번호 부여 메소드
    def get_scr_no(self):
        if self.g_scr_no < 9999:
            self.g_scr_no += 1
        else:
            self.g_scr_no = 1000

        return str(self.g_scr_no)

    def get_cur_price(self, l_jongmok_cd):
        l_for_flag = 0
        self.g_cur_price2 = sys.maxsize

        l_scr_no = self.get_scr_no()

        # while True:
        self.g_rqname = '현재가조회'
        # self.g_flag_6 = 1
        self.kiwoom.dynamicCall("SetInputValue(QString, QString)", ['종목코드', l_jongmok_cd])

        # l_scr_no = self.get_scr_no()

        # 현재가 조회 요청
        self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)",
                                [self.g_rqname, 'opt10001', 0, l_scr_no])
        # try:
        #     l_for_cnt = 0
        #     while True:
        #         if self.g_flag_6 == 1:
        #             time.sleep(0.2)
        #             self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
        #             l_for_flag = 1
        #             break
        #         else:
        #             time.sleep(0.2)
        #             l_for_cnt += 1
        #             print(self.g_flag_6)
        #             print(l_for_cnt)
        #             if l_for_cnt == 5:
        #                 l_for_flag = 0
        #                 break
        #             else:
        #                 continue
        # except Exception as ex:
        #     print(str(ex))
        # if l_for_flag == 1:
        #     break
        # elif l_for_flag == 0:
        #     time.sleep(0.2)
        #     continue
        # time.sleep(0.2)
        time.sleep(0.5)

        self.kiwoom.dynamicCall("DisconnectRealData(QString)", l_scr_no)
        # return self.g_cur_price2

        # if l_for_flag == 1:
        #     break
        # elif l_for_flag == 0:
        #     time.sleep(0.2)
        #     continue
        # time.sleep(0.2)
