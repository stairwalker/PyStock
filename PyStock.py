import sys
from datetime import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QAxContainer import *
from Kiwoom import *
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
        self.kiwoom = Kiwoom()
        self.kiwoom.comm_connect() 

	    #키움 API 신호
        #self.kiwoom.OnEventConnect.connect(self._event_connect)
        #self.kiwoom.OnReceiveTrData.connect(self._receive_tr_data)

        # 타이틀 설정
        self.setWindowTitle("종목 코드")
        self.setGeometry(300, 300, 300, 180)

        #1초 타이머 설정
        self.timer1s = QTimer(self)
        self.timer1s.start(1000)
        self.timer1s.timeout.connect(self.timer1sec)

        # 400ms 타이머 설정
        self.timer10ms = QTimer(self)
        self.timer10ms.start(400)
        self.timer10ms.timeout.connect(self.getdatatimer)
        self.getdataflag = 0
        self.getdatacount = -1

        self.prvFilename = ""

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

        state = self.kiwoom.get_connect_state()
        if state == 1:
            state_msg = "서버 연결 성공"
            self.btn1.setDisabled(1)
            self.btn2.setDisabled(0)
            self.btn3.setDisabled(0)
        else:
            state_msg = "서버 미접속"
            self.btn1.setDisabled(0)
            self.btn2.setDisabled(1)
            self.btn3.setDisabled(1)
        self.statusbar.showMessage(state_msg + " | " + time_msg)



    def getdatatimer(self):
        if self.kiwoom.dataReceived == 1:
            current_date = QDate.currentDate()            
            filename = '..\\data\\' + current_date.toString("yyyyMMdd") + '_' + self.codelist[self.getdatacount] + '.csv'
            
            if filename == self.prvFilename :                                
                with open(filename,'a') as sfile:
                    i=0
                    if self.kiwoom.dataunit != None :                    
                        while i<len(self.kiwoom.dataunit) :
                            for j in range(0,6):
                                if j==2 : # 날짜인 경우
                                    sfile.write("'"+self.kiwoom.dataunit[i][j]+"'")
                                else :
                                    sfile.write(self.kiwoom.dataunit[i][j])
                                if j != 5:                                    
                                    sfile.write(",")
                                else:
                                    sfile.write('\n')
                            i=i+1                                    
                        self.kiwoom.dataReceived = 0                
            else:
                if self.prvFilename != "" :
                    time.sleep(12)
                logfilename = 'lastest.txt'
                with open(logfilename,'w') as logfile:
                    logfile.write(self.prvFilename)
                current_time = QTime.currentTime()
                self.listWidget.insertItem(0,current_time.toString("hh:mm:ss") +" - "+ self.codelist[self.getdatacount] + " 조회 시작")
                self.prvFilename = filename
                with open(filename,'w') as sfile:
                    i=0
                    if self.kiwoom.dataunit != None :                    
                        while i<len(self.kiwoom.dataunit) :
                            for j in range(0,6):
                                if j==2 : # 날짜인 경우
                                    sfile.write("'"+self.kiwoom.dataunit[i][j]+"'")
                                else :
                                    sfile.write(self.kiwoom.dataunit[i][j])

                                if j != 5:
                                    sfile.write(',')
                                else:
                                    sfile.write('\n')
                            i=i+1                                    
                        self.kiwoom.dataReceived = 0

        else :            
            if self.getdataflag == 1:
                if self.kiwoom.remained_data == False : 
                    self.getdatacount = self.getdatacount + 1

                if self.getdatacount < len(self.codelist) :
                    self.kiwoom.set_input_value("종목코드", self.codelist[self.getdatacount])
                    self.kiwoom.set_input_value("시작일자", "20170805")
                    self.kiwoom.set_input_value("종료일자", "20170805")
                    self.kiwoom.set_input_value("수정주가구분", 0)
                    if self.kiwoom.remained_data == True : 
                        self.kiwoom.comm_rq_data("주식분봉차트조회", "opt10080", 2, "1001")                        
                    else :
                        self.kiwoom.comm_rq_data("주식분봉차트조회", "opt10080", 0, "1001")                
                else:
                    self.getdataflag = 0
                    self.getdatacount = 0

    def btn1_clicked(self):        
        self.kiwoom.comm_connect()
        print(ret)
        if ret == 0 :
            self.listWidget.insertItem(1,"로그인 성공")
        else :
            self.listWidget.insertItem(1, "로그오프")

    def btn2_clicked(self):
        codestring = self.kiwoom.get_kospi_code_list()
        self.codelist = codestring.split(";")
        self.getdataflag = 1

    def btn3_clicked(self):
        files = glob.glob('..\\data\\*.csv')
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


if __name__ == "__main__":

    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    sys.exit(app.exec_())