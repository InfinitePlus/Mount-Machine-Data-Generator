import sys
import os
from PyQt5.QtWidgets import *
import csv
import random
import winreg

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
    return winreg.QueryValueEx(key, "Desktop")[0]
#C:\Users\liu.peng56\Desktop
class Winform(QWidget):
    def __init__(self, parent=None):
        super(Winform, self).__init__(parent)
        self.initUI()

    def initUI(self):
        vlayout = QVBoxLayout()

        hlayout1= QHBoxLayout()
        label_liehao=QLabel("列号:")
        #liehao=QLineEdit("P1")
        self.liehao=QComboBox()
        self.liehao.addItems(["P1","P2"])
        hlayout1.addWidget(label_liehao, 0)
        hlayout1.addWidget(self.liehao, 1)
        vlayout.addLayout(hlayout1)

        hlayout2 = QHBoxLayout()
        label_mzh = QLabel("模组号:") # 模组号
        self.mzh = QComboBox()
        self.mzh.addItems(["M3IIIX", "M6IIIX", "M3IIISX", "M6IIISX"])
        hlayout2.addWidget(label_mzh, 0)
        hlayout2.addWidget(self.mzh, 1)
        vlayout.addLayout(hlayout2)

        hlayout3 = QHBoxLayout()
        label_mzxlh = QLabel("模组序列号:No")  # 模组序列号
        #self.mzxlh = QComboBox()
        #self.mzxlh.addItems(["No4319","No21930","No26858","No13201","No15444"])
        self.mzxlh=QLineEdit("4319")
        hlayout3.addWidget(label_mzxlh, 0)
        hlayout3.addWidget(self.mzxlh, 1)
        vlayout.addLayout(hlayout3)

        hlayout4 = QHBoxLayout()
        label_gztlx = QLabel("工作头类型:")  # 工作头类型
        self.gztlx = QComboBox()
        self.gztlx.addItems(["H01","H02F","H04SF","H24","V12"])
        hlayout4.addWidget(label_gztlx, 0)
        hlayout4.addWidget(self.gztlx, 1)
        vlayout.addLayout(hlayout4)

        hlayout5 = QHBoxLayout()
        label_mzxlh2 = QLabel("模组序列号2:")  # 模组序列号2
        #self.mzxlh2 = QComboBox()
        #self.mzxlh2.addItems(["037740", "021021", "006189", "008627", "006320"])
        self.mzxlh2 = QLineEdit("037740")
        hlayout5.addWidget(label_mzxlh2, 0)
        hlayout5.addWidget(self.mzxlh2, 1)
        vlayout.addLayout(hlayout5)

        hlayout6 = QHBoxLayout()
        label_hz = QLabel("后缀:")  # 后缀
        self.hz = QComboBox()
        self.hz.addItems(["HBCC", "HBCG"])
        hlayout6.addWidget(label_hz, 0)
        hlayout6.addWidget(self.hz, 1)
        vlayout.addLayout(hlayout6)
        #状态栏
        self.status=QLabel()
        vlayout.addWidget(self.status)
        #按钮
        startbutton=QPushButton("开始生成")
        startbutton.clicked.connect(self.CreateCsv)
        vlayout.addWidget(startbutton)

        self.setLayout(vlayout)
    def CreateCsv(self):
        if(self.gztlx.currentText()=="H01"):
            try:
                with open('P1_M6_M6IIIXNo15444_H01_006320_HBCG.csv', 'r') as f:
                    reader = csv.reader(f)
                    result = list(reader)
                    if(self.mzh.currentText()=="M3IIIX"):
                        filename = "{}_M3_M3IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(),"No"+self.mzxlh.text(),self.gztlx.currentText(), self.mzxlh2.text(),self.hz.currentText())
                        result[0][0] = "M3\tM3IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format("No"+self.mzxlh.text(),self.gztlx.currentText(),self.mzxlh2.text())
                    elif(self.mzh.currentText()=="M6IIIX"):
                        filename = "{}_M6_M6IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(),"No"+self.mzxlh.text(),self.gztlx.currentText(),self.mzxlh2.text(),self.hz.currentText())
                        result[0][0] = "M6\tM6IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format("No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M3IIISX"):
                        filename = "{}_M3_M3IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(),"No"+self.mzxlh.text(),self.gztlx.currentText(), self.mzxlh2.text(),self.hz.currentText())
                        result[0][0] = "M3\tM3IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format("No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIISX"):
                        filename = "{}_M6_M6IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(),"No"+self.mzxlh.text(),self.gztlx.currentText(),self.mzxlh2.text(),self.hz.currentText())
                        result[0][0] = "M6\tM6IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format("No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    for i in range(1, len(result)):
                        x = result[i][0].split("\t", 6)

                        x[4] = x[4].strip()
                        x4max=float(x[4]) + 3
                        x4min = float(x[4]) - 3
                        if (x4min > 10):
                            x4min=8
                            x4max=10
                        if (x4max > 10):
                            x4max=10
                        if (x4max < -10):
                            x4max=-8
                            x4min=-10
                        if (x4min < -10):
                            x4min=-10
                        x4num=float(format(random.uniform(x4min, x4max),'.6f'))
                        if(x4num>0):
                            x[4] = "  "+str(x4num)
                        else:
                            x[4] = " " +str(x4num)
                        j=0
                        while(j<(6-len(str(x4num).split(".")[1]))):
                            x[4]=x[4]+"0"
                            j=j+1

                        x[5] = x[5].strip()
                        x5max = float(x[4]) + 3
                        x5min = float(x[4]) - 3
                        if (x5min > 10):
                            x5min=8
                            x5max=10
                        if (x5max > 10):
                            x5max=10
                        if (x5max < -10):
                            x5max=-8
                            x5min=-10
                        if (x5min < -10):
                            x5min=-10
                        x5num=float(format(random.uniform(x5min, x5max),'.6f'))
                        if (x5num > 0):
                            x[5] = "  " +str(x5num)
                        else:
                            x[5] = " " + str(x5num)
                        j = 0
                        while (j < (6 - len(str(x5num).split(".")[1]))):
                            x[5] = x[5] + "0"
                            j = j + 1

                        x[6] = x[6].strip()
                        x6max=float(x[6]) + 0.2
                        x6min=float(x[6]) - 0.2
                        if (x6min > 0.5):
                            x6min = 0
                            x6max = 0.4
                        if (x6max > 0.5):
                            x6max = 0.5
                        if (x6max < -0.5):
                            x6max = -0.4
                            x6min = 0
                        if (x6min < -0.5):
                            x6min = -0.5
                        x6num=float(format(random.uniform(x6min,x6max),'.6f'))

                        if (x6num > 0):
                            x[6] = "  " +str(x6num)
                        else:
                            x[6] = " " + str(x6num)
                        j = 0
                        while (j < (6 - len(str(x6num).split(".")[1]))):
                            x[6] = x[6] + "0"
                            j = j + 1
                        tsym = "\t"
                        result[i][0] = x[0] + tsym +  x[1] + tsym + x[2] + tsym + x[3] + tsym + x[4] + tsym + x[5] + tsym + x[6]
            except Exception:
                self.status.setText("找不到模板文件 P1_M6_M6IIIXNo15444_H01_006320_HBCG.csv")
            else:
                self.status.setText("文件生成成功")
        elif(self.gztlx.currentText()=="H02F"):
            try:
                with open('P1_M6_M6IIIXNo13201_H02F_008627_HBCG.csv', 'r') as f:
                    reader = csv.reader(f)
                    result = list(reader)
                    if (self.mzh.currentText() == "M3IIIX"):
                        filename = "{}_M3_M3IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M3\tM3IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIIX"):
                        filename = "{}_M6_M6IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M6\tM6IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M3IIISX"):
                        filename = "{}_M3_M3IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M3\tM3IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIISX"):
                        filename = "{}_M6_M6IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M6\tM6IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    for i in range(1, len(result)):
                        x = result[i][0].split("\t", 6)

                        x[4] = x[4].strip()
                        x4max = float(x[4]) + 3
                        x4min = float(x[4]) - 3
                        if (x4min > 10):
                            x4min = 8
                            x4max = 10
                        if (x4max > 10):
                            x4max = 10
                        if (x4max < -10):
                            x4max = -8
                            x4min = -10
                        if (x4min < -10):
                            x4min = -10
                        x4num = float(format(random.uniform(x4min, x4max), '.6f'))
                        if (x4num > 0):
                            x[4] = "  " + str(x4num)
                        else:
                            x[4] = " " + str(x4num)
                        j = 0
                        while (j < (6 - len(str(x4num).split(".")[1]))):
                            x[4] = x[4] + "0"
                            j = j + 1

                        x[5] = x[5].strip()
                        x5max = float(x[4]) + 3
                        x5min = float(x[4]) - 3
                        if (x5min > 10):
                            x5min = 8
                            x5max = 10
                        if (x5max > 10):
                            x5max = 10
                        if (x5max < -10):
                            x5max = -8
                            x5min = -10
                        if (x5min < -10):
                            x5min = -10
                        x5num = float(format(random.uniform(x5min, x5max), '.6f'))
                        if (x5num > 0):
                            x[5] = "  " + str(x5num)
                        else:
                            x[5] = " " + str(x5num)
                        j = 0
                        while (j < (6 - len(str(x5num).split(".")[1]))):
                            x[5] = x[5] + "0"
                            j = j + 1

                        x[6] = x[6].strip()
                        x6max = float(x[6]) + 0.2
                        x6min = float(x[6]) - 0.2
                        if (x6min > 0.5):
                            x6min = 0
                            x6max = 0.4
                        if (x6max > 0.5):
                            x6max = 0.5
                        if (x6max < -0.5):
                            x6max = -0.4
                            x6min = 0
                        if (x6min < -0.5):
                            x6min = -0.5
                        x6num = float(format(random.uniform(x6min, x6max), '.6f'))
                        if (x6num > 0):
                            x[6] = "  " + str(x6num)
                        else:
                            x[6] = " " + str(x6num)
                        j = 0
                        while (j < (6 - len(str(x6num).split(".")[1]))):
                            x[6] = x[6] + "0"
                            j = j + 1
                        tsym = "\t"
                        result[i][0] = x[0] + tsym + x[1] + tsym + x[2] + tsym + x[3] + tsym + x[4] + tsym + x[5] + tsym + x[6]
            except Exception:
                self.status.setText("找不到模板文件 P1_M6_M6IIIXNo13201_H02F_008627_HBCG.csv")
            else:
                self.status.setText("文件生成成功")
        elif (self.gztlx.currentText() == "H04SF"):
            try:
                with open('P1_M3_M3IIIXNo26858_H04SF_006189_HBCG.csv','r') as f:
                    reader = csv.reader(f)
                    result = list(reader)
                    if (self.mzh.currentText() == "M3IIIX"):
                        filename = "{}_M3_M3IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M3\tM3IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIIX"):
                        filename = "{}_M6_M6IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M6\tM6IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M3IIISX"):
                        filename = "{}_M3_M3IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M3\tM3IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIISX"):
                        filename = "{}_M6_M6IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M6\tM6IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    for i in range(1, len(result)):
                        x = result[i][0].split("\t", 6)

                        x[4] = x[4].strip()
                        x4max = float(x[4]) + 3
                        x4min = float(x[4]) - 3
                        if (x4min > 10):
                            x4min = 8
                            x4max = 10
                        if (x4max > 10):
                            x4max = 10
                        if (x4max < -10):
                            x4max = -8
                            x4min = -10
                        if (x4min < -10):
                            x4min = -10
                        x4num = float(format(random.uniform(x4min, x4max), '.6f'))
                        if (x4num > 0):
                            x[4] = "  " + str(x4num)
                        else:
                            x[4] = " " + str(x4num)
                        j = 0
                        while (j < (6 - len(str(x4num).split(".")[1]))):
                            x[4] = x[4] + "0"
                            j = j + 1

                        x[5] = x[5].strip()
                        x5max = float(x[4]) + 3
                        x5min = float(x[4]) - 3
                        if (x5min > 10):
                            x5min = 8
                            x5max = 10
                        if (x5max > 10):
                            x5max = 10
                        if (x5max < -10):
                            x5max = -8
                            x5min = -10
                        if (x5min < -10):
                            x5min = -10
                        x5num = float(format(random.uniform(x5min, x5max), '.6f'))
                        if (x5num > 0):
                            x[5] = "  " + str(x5num)
                        else:
                            x[5] = " " + str(x5num)
                        j = 0
                        while (j < (6 - len(str(x5num).split(".")[1]))):
                            x[5] = x[5] + "0"
                            j = j + 1

                        x[6] = x[6].strip()
                        x6max = float(x[6]) + 0.2
                        x6min = float(x[6]) - 0.2
                        if (x6min > 0.5):
                            x6min = 0
                            x6max = 0.4
                        if (x6max > 0.5):
                            x6max = 0.5
                        if (x6max < -0.5):
                            x6max = -0.4
                            x6min = 0
                        if (x6min < -0.5):
                            x6min = -0.5
                        x6num = float(format(random.uniform(x6min, x6max), '.6f'))
                        if (x6num > 0):
                            x[6] = "  " + str(x6num)
                        else:
                            x[6] = " " + str(x6num)
                        j = 0
                        while (j < (6 - len(str(x6num).split(".")[1]))):
                            x[6] = x[6] + "0"
                            j = j + 1
                        tsym = "\t"
                        result[i][0] = x[0] + tsym + x[1] + tsym + x[2] + tsym + x[3] + tsym + x[4] + tsym + x[5] + tsym + x[6]
            except Exception:
                self.status.setText("找不到模板文件 P1_M3_M3IIIXNo26858_H04SF_006189_HBCG.csv")
            else:
                self.status.setText("文件生成成功")
        elif (self.gztlx.currentText() == "H24"):
            try:
                with open('P1_M3_M3IIISXNo4319_H24_037740_HBCC.csv','r') as f:
                    reader = csv.reader(f)
                    result = list(reader)
                    if (self.mzh.currentText() == "M3IIIX"):
                        filename = "{}_M3_M3IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M3\tM3IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIIX"):
                        filename = "{}_M6_M6IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M6\tM6IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M3IIISX"):
                        filename = "{}_M3_M3IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M3\tM3IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIISX"):
                        filename = "{}_M6_M6IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M6\tM6IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    for i in range(1, len(result)):
                        x = result[i][0].split("\t", 6)
                        x[4] = x[4].strip()
                        x4max = float(x[4]) + 3
                        x4min = float(x[4]) - 3
                        if (x4min > 10):
                            x4min = 8
                            x4max = 10
                        if (x4max > 10):
                            x4max = 10
                        if (x4max < -10):
                            x4max = -8
                            x4min = -10
                        if (x4min < -10):
                            x4min = -10
                        x4num = float(format(random.uniform(x4min, x4max), '.6f'))
                        if (x4num > 0):
                            x[4] = "  " + str(x4num)
                        else:
                            x[4] = " " + str(x4num)
                        j = 0
                        while (j < (6 - len(str(x4num).split(".")[1]))):
                            x[4] = x[4] + "0"
                            j = j + 1

                        x[5] = x[5].strip()
                        x5max = float(x[4]) + 3
                        x5min = float(x[4]) - 3
                        if (x5min > 10):
                            x5min = 8
                            x5max = 10
                        if (x5max > 10):
                            x5max = 10
                        if (x5max < -10):
                            x5max = -8
                            x5min = -10
                        if (x5min < -10):
                            x5min = -10
                        x5num = float(format(random.uniform(x5min, x5max), '.6f'))
                        if (x5num > 0):
                            x[5] = "  " + str(x5num)
                        else:
                            x[5] = " " + str(x5num)
                        j = 0
                        while (j < (6 - len(str(x5num).split(".")[1]))):
                            x[5] = x[5] + "0"
                            j = j + 1

                        x[6] = x[6].strip()
                        x6max = float(x[6]) + 0.2
                        x6min = float(x[6]) - 0.2
                        if (x6min > 0.5):
                            x6min = 0
                            x6max = 0.4
                        if (x6max > 0.5):
                            x6max = 0.5
                        if (x6max < -0.5):
                            x6max = -0.4
                            x6min = 0
                        if (x6min < -0.5):
                            x6min = -0.5
                        x6num = float(format(random.uniform(x6min, x6max), '.6f'))
                        if (x6num > 0):
                            x[6] = "  " + str(x6num)
                        else:
                            x[6] = " " + str(x6num)
                        j = 0
                        while (j < (6 - len(str(x6num).split(".")[1]))):
                            x[6] = x[6] + "0"
                            j = j + 1
                        tsym = "\t"
                        result[i][0] = x[0] + tsym + x[1] + tsym + x[2] + tsym + x[3] + tsym + x[4] + tsym + x[5] + tsym + x[6]
            except Exception:
                self.status.setText("找不到模板文件 P1_M3_M3IIISXNo4319_H24_037740_HBCC.csv")
            else:
                self.status.setText("文件生成成功")
        elif (self.gztlx.currentText() == "V12"):
            try:
                with open('P1_M3_M3IIISXNo21930_V12_021021_HBCC.csv', 'r') as f:
                    reader = csv.reader(f)
                    result = list(reader)
                    if (self.mzh.currentText() == "M3IIIX"):
                        filename = "{}_M3_M3IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M3\tM3IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIIX"):
                        filename = "{}_M6_M6IIIX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                        self.gztlx.currentText(), self.mzxlh2.text(),
                                                                        self.hz.currentText())
                        result[0][0] = "M6\tM6IIIX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M3IIISX"):
                        filename = "{}_M3_M3IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M3\tM3IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    elif (self.mzh.currentText() == "M6IIISX"):
                        filename = "{}_M6_M6IIISX{}_{}_{}_{}.csv".format(self.liehao.currentText(), "No"+self.mzxlh.text(),
                                                                         self.gztlx.currentText(), self.mzxlh2.text(),
                                                                         self.hz.currentText())
                        result[0][0] = "M6\tM6IIISX{}\t{}\t{}\t0\t\t\t0\t0\t0\t1\t8\t0\t0\t\t\t0\t".format(
                            "No"+self.mzxlh.text(), self.gztlx.currentText(), self.mzxlh2.text())
                    for i in range(1, len(result)):
                        x = result[i][0].split("\t", 6)
                        x[4] = x[4].strip()
                        x4max = float(x[4]) + 3
                        x4min = float(x[4]) - 3
                        if (x4min > 10):
                            x4min = 8
                            x4max = 10
                        if (x4max > 10):
                            x4max = 10
                        if (x4max < -10):
                            x4max = -8
                            x4min = -10
                        if (x4min < -10):
                            x4min = -10
                        x4num = float(format(random.uniform(x4min, x4max), '.6f'))
                        if (x4num > 0):
                            x[4] = "  " + str(x4num)
                        else:
                            x[4] = " " + str(x4num)
                        x[5] = x[5].strip()
                        x5max = float(x[4]) + 3
                        x5min = float(x[4]) - 3
                        if (x5min > 10):
                            x5min = 8
                            x5max = 10
                        if (x5max > 10):
                            x5max = 10
                        if (x5max < -10):
                            x5max = -8
                            x5min = -10
                        if (x5min < -10):
                            x5min = -10
                        x5num = float(format(random.uniform(x5min, x5max), '.6f'))
                        if (x5num > 0):
                            x[5] = "  " + str(x5num)
                        else:
                            x[5] = " " + str(x5num)
                        x[6] = x[6].strip()
                        x6max = float(x[6]) + 0.2
                        x6min = float(x[6]) - 0.2
                        if (x6min > 0.5):
                            x6min = 0
                            x6max = 0.4
                        if (x6max > 0.5):
                            x6max = 0.5
                        if (x6max < -0.5):
                            x6max = -0.4
                            x6min = 0
                        if (x6min < -0.5):
                            x6min = -0.5
                        x6num = float(format(random.uniform(x6min, x6max), '.6f'))
                        if (x6num > 0):
                            x[5] = "  " + str(x6num)
                        else:
                            x[5] = " " + str(x6num)
                        tsym = "\t"
                        bksym = " "
                        result[i][0] = x[0] + tsym + " " + x[1] + tsym + x[2] + tsym + "  " + x[3] + tsym + x[4] + tsym + x[5] + tsym + x[6]
            except Exception:
                self.status.setText("找不到模板文件 P1_M3_M3IIISXNo21930_V12_021021_HBCC.csv")
            else:
                self.status.setText("文件生成成功")
        else:
            print("read error")

        try:
            with open(get_desktop()+"\\"+filename, 'w', encoding='utf-8', newline='') as f2:
                writer = csv.writer(f2)
                for item in result:
                    writer.writerow(item)
                print("ok")
        except Exception:
            print("join error")
        # 重命名
        try:
            #Path = get_desktop()
            Olddir = get_desktop()+"\\"+filename
            qf=os.path.splitext(filename)
            Newdir = get_desktop()+"\\"+qf[0]+".dat"
            if os.path.exists(Newdir):  # 如果文件存在
                # 删除文件
                os.remove(Newdir)
            os.rename(Olddir, Newdir)
        except Exception:
            self.status.setText("重命名失败")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    form = Winform()
    form.setWindowTitle("贴片机文件生成")
    form.setMinimumWidth(300)
    form.show()
    sys.exit(app.exec_())