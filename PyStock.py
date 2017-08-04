import sys
from datetime import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *
import os
import glob


def adjusttime() :
    return 1

def endtime() :
    return 15

class MyWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        #QtGui.QMainWindow.__init__(self,parent)
        #self.setupUi(self)

        # 키움 API 접속
        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")


        # 타이틀 설정
        self.setWindowTitle("종목 코드")
        self.setGeometry(300, 300, 300, 180)

        #1초 타이머 설정
        self.timer1s = QTimer(self)
        self.timer1s.start(1000)
        self.timer1s.timeout.connect(self.timer1sec)

        # 1초 타이머 설정
        self.timer10ms = QTimer(self)
        self.timer10ms.start(400)
        self.timer10ms.timeout.connect(self.getdatatimer)
        self.getdataflag = 0
        self.getdatacount = 0

        self.temp = 0

        self.setupUI()

    def setupUI(self):

        # StatusBar
        self.statusbar = QStatusBar(self)
        self.setStatusBar(self.statusbar)

        # button 1
        self.btn1 = QPushButton("연결", self)
        self.btn1.move(190, 10)
        self.btn1.clicked.connect(self.btn1_clicked)

        # button 2
        self.btn2 = QPushButton("데이터 산출", self)
        self.btn2.move(190, 45)
        self.btn2.clicked.connect(self.btn2_clicked)

        # button 3
        self.btn3 = QPushButton("데이터 분절", self)
        self.btn3.move(190, 80)
        self.btn3.clicked.connect(self.btn3_clicked)

        self.btn1.setDisabled(1)
        self.btn2.setDisabled(1)
        self.btn3.setDisabled(1)

        self.listWidget = QListWidget(self)
        self.listWidget.setGeometry(10, 10, 170, 130)

    def timer1sec(self):
        current_time = QTime.currentTime()
        text_time = current_time.toString("hh:mm:ss")
        time_msg = "현재시간: " + text_time

        state = self.kiwoom.dynamicCall("GetConnectState()")
        if state == 1:
            state_msg = "서버 연결 성공"
            self.btn1.setDisabled(1)
            self.btn2.setDisabled(0)
            self.btn3.setDisabled(0)
            #self.kiwoom.dynamicCall("CommConnect()")
        else:
            state_msg = "서버 미접속"
            self.btn1.setDisabled(0)
            self.btn2.setDisabled(1)
            self.btn3.setDisabled(1)





        self.statusbar.showMessage(state_msg + " | " + time_msg)

    def getdatatimer(self):
        if self.getdataflag == 1:
            if self.getdatacount < 1231:
                ret = self.kiwoom.dynamicCall("SetInputValue(QString,QString)", "종목코드", self.codelist[self.getdatacount])
                ret = self.kiwoom.dynamicCall("CommRqData(Qstring,Qstring,Qstring,QString)", "주식분봉차트조회","OPT10080","0",self.codelist[self.getdatacount])
                self.getdatacount = self.getdatacount + 1
            else:
                self.getdataflag = 0
                self.getdatacount = 0




    def btn1_clicked(self):
        #stockCode = "141000"
        #ret = self.kiwoom.dynamicCall("SetInputValue(QString,QString)","종목코드",stockCode)
        ret = self.kiwoom.dynamicCall("CommConnect()")
        print(ret)
        if ret == 0 :
            self.listWidget.insertItem(1,"Connected")
        else :
            self.listWidget.insertItem(1, "Not connected")



    def btn2_clicked(self):
        codestring = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ["0"])
        self.codelist = codestring.split(";")
        self.getdataflag = 1

    def btn3_clicked(self):
        files = glob.glob('.\\data\\*.csv')
        i=0
        while len(files) > i:
            with open(files[i], 'r') as sfile:    #파일 순차적으로 열기
                fileline = sfile.readlines()
                linecount=0
                filecount=0
                oclock = 0
                while len(fileline) > linecount:
                    featuredata = fileline[linecount].split(",")
                    if oclock != int(featuredata[2][9:10]) :     #시간이 변경되면 파일 변경
                        oclock = int(featuredata[2][9:10])

                    with open(files[i], 'w') as sfile:
                        i=i+1



    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        #분봉차트 데이터 수신시
        if sRQName == "주식분봉차트조회" :
            revdata = self.kiwoom.dynamicCall("GetCommDataEx(QString,QString)", "OPT10080", "주식분봉차트조회")
            self.temp = self.temp + 1
            print(self.temp)
            filename = '.\\data\\' + revdata[0][2][:8] + '_' + sScrNo + '.csv'
            with open(filename,'w') as sfile:
                i=0
                while revdata[0][2][:8] == revdata[i][2][:8]:
                    for j in range(0,6):
                        sfile.write(revdata[i][j])
                        if j != 5:
                            sfile.write(',')
                    sfile.write('\n')
                    i=i+1

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    sys.exit(app.exec_())