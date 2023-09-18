# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt 
import pandas as pd
import numpy as np 

from scipy.optimize import curve_fit
import sys
import time

import pyqtgraph as pg
from PyQt5 import QtCore, QtGui, QtWidgets 
from PyQt5.QtWidgets import QFileDialog
import os  



class Ui_Dialog(QtWidgets.QDialog):
    file_path = ""
    def setupUi(self, Dialog): 
         
        Dialog.setObjectName("Dialog")
        Dialog.resize(1280,800)

        self.gridLayoutWidget = QtWidgets.QWidget(Dialog)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 431, 100))
        self.gridLayoutWidget.setObjectName("gridLayoutWidget")

        self.gridLayout = QtWidgets.QGridLayout(self.gridLayoutWidget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")

        self.offsetOfPoly = QtWidgets.QDoubleSpinBox(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.offsetOfPoly.setFont(font)
        self.offsetOfPoly.setObjectName("offsetOfPoly")
        self.offsetOfPoly.setProperty("value", float(60))
        self.offsetOfPoly.setMaximum(1000.0)
        self.gridLayout.addWidget(self.offsetOfPoly, 1, 1, 1, 1)

        self.label_2 = QtWidgets.QLabel(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_2.setFont(font)
        self.label_2.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_2.setObjectName("label_2")
        self.gridLayout.addWidget(self.label_2, 1, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_3.setFont(font)
        self.label_3.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_3.setObjectName("label_3")
        self.gridLayout.addWidget(self.label_3, 2, 0, 1, 1)
        self.label = QtWidgets.QLabel(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label.setFont(font)
        self.label.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label.setObjectName("label")
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.littlePolyDegree = QtWidgets.QSpinBox(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.littlePolyDegree.setFont(font)
        self.littlePolyDegree.setMaximum(1000)
        self.littlePolyDegree.setObjectName("littlePolyDegree")
        
        self.littlePolyDegree.setProperty("value", 3)
        self.gridLayout.addWidget(self.littlePolyDegree, 2, 1, 1, 1)
        self.bigPolyDegree = QtWidgets.QSpinBox(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.bigPolyDegree.setFont(font)
        self.bigPolyDegree.setObjectName("bigPolyDegree")
        self.bigPolyDegree.setProperty("value", 9)
        self.bigPolyDegree.setMaximum(1000)
        self.gridLayout.addWidget(self.bigPolyDegree, 0, 1, 1, 1)

        font = QtGui.QFont()
        font.setPointSize(14)
        self.OK = QtWidgets.QPushButton(Dialog)
        self.OK.setGeometry(QtCore.QRect(900,750, 300, 30))
        self.OK.setFont(font)

        self.progressBar = QtWidgets.QProgressBar(Dialog)
        self.progressBar.setGeometry(QtCore.QRect(10, 755, 800, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar") 
        self.progressBar.setFont(font)

        self.fileLocate = QtWidgets.QPushButton(Dialog)
        self.fileLocate.setGeometry(QtCore.QRect(500, 40, 351, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.fileLocate.setFont(font)
        self.fileLocate.setObjectName("fileLocate")

        self.filePathEdit = QtWidgets.QLineEdit(Dialog)
        self.filePathEdit.setGeometry(QtCore.QRect(500, 10, 351, 21))
        self.filePathEdit.setObjectName("filePathEdit")

        self.windowsSize = QtWidgets.QSpinBox(Dialog)
        font = QtGui.QFont()
        font.setPointSize(13)
        self.windowsSize.setFont(font)
        self.windowsSize.setObjectName("windowsSize")
        self.windowsSize.setMaximum(1000)
        self.windowsSize.setProperty("value", 0)
        self.windowsSize.setGeometry(QtCore.QRect(750, 80, 100, 30))
        self.windowsSizeLabel = QtWidgets.QLabel(Dialog)
        self.windowsSizeLabel.setFont(font)
        self.windowsSizeLabel.setObjectName("windowsSizeLabel")
        self.windowsSizeLabel.setGeometry(QtCore.QRect(500, 80, 250, 30))

        self.horizontalLayoutWidget = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(10,140, 1270, 50))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_4 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout.addWidget(self.label_4)

        self.xmaxGraph = QtWidgets.QDoubleSpinBox(self.horizontalLayoutWidget)
        self.xmaxGraph.setMaximum(100000.0)
        self.xmaxGraph.setFont(font)
        self.xmaxGraph.setObjectName("maxGraph")
        self.horizontalLayout.addWidget(self.xmaxGraph)

        self.label_5 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout.addWidget(self.label_5)
        self.xminGraph = QtWidgets.QDoubleSpinBox(self.horizontalLayoutWidget)

        
        self.xminGraph.setFont(font)
        self.xminGraph.setObjectName("minGraph")
        self.horizontalLayout.addWidget(self.xminGraph)     
        
        self.label_6 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label_6.setFont(font)
        self.label_6.setObjectName("label_6")
        self.horizontalLayout.addWidget(self.label_6)
        self.xminGraph.setMaximum(100000.0)
        
        self.ymaxGraph = QtWidgets.QDoubleSpinBox(self.horizontalLayoutWidget)
        self.ymaxGraph.setMaximum(100000.0)
        self.ymaxGraph.setFont(font)
        self.ymaxGraph.setObjectName("ymaxGraph")
        self.horizontalLayout.addWidget(self.ymaxGraph)

        self.label_7 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        self.label_7.setFont(font)
        self.label_7.setObjectName("label_7")
        self.horizontalLayout.addWidget(self.label_7)
        self.xminGraph.setMaximum(100000.0)
        
        self.yminGraph = QtWidgets.QDoubleSpinBox(self.horizontalLayoutWidget)
        self.yminGraph.setMaximum(100000.0)
        self.yminGraph.setFont(font)
        self.yminGraph.setObjectName("yminGraph")
        self.horizontalLayout.addWidget(self.yminGraph)

        self.regenerateGraph = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.regenerateGraph.setFont(font)
        self.regenerateGraph.setObjectName("regenerateGraph")
        self.horizontalLayout.addWidget(self.regenerateGraph)

        self.line_2 = QtWidgets.QFrame(Dialog)
        self.line_2.setGeometry(QtCore.QRect(0, 135, 1280, 20))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")

        self.activateSecondPol = QtWidgets.QCheckBox(Dialog)
        self.activateSecondPol.setGeometry(QtCore.QRect(20, 112, 331, 25))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.activateSecondPol.setFont(font)
        self.activateSecondPol.setObjectName("activateSecondPol")

        self.activateSecondLorentzian = QtWidgets.QCheckBox(Dialog)
        self.activateSecondLorentzian.setGeometry(QtCore.QRect(250, 112, 331, 25))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.activateSecondLorentzian.setFont(font)
        self.activateSecondLorentzian.setObjectName("activateSecondLorentzian")

        self.activateFirstLorentzian = QtWidgets.QCheckBox(Dialog)
        self.activateFirstLorentzian.setGeometry(QtCore.QRect(480, 112, 331, 25))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.activateFirstLorentzian.setFont(font)
        self.activateFirstLorentzian.setObjectName("activateFirstLorentzian")


        self.bigPolyOut = QtWidgets.QCheckBox(Dialog)
        self.bigPolyOut.setGeometry(QtCore.QRect(900, 10, 331, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.bigPolyOut.setFont(font)
        self.bigPolyOut.setObjectName("bigPolyOut")

        self.littlePolyOut = QtWidgets.QCheckBox(Dialog)
        self.littlePolyOut.setGeometry(QtCore.QRect(900, 70, 331, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.littlePolyOut.setFont(font)
        self.littlePolyOut.setObjectName("littlePolyOut")
        self.littlePolyOut.setEnabled(False)

        self.peaksOut = QtWidgets.QCheckBox(Dialog)
        self.peaksOut.setGeometry(QtCore.QRect(900, 40, 331, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.peaksOut.setFont(font)
        self.peaksOut.setObjectName("peaksOut")

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

        self.fileLocate.clicked.connect(self.browsefiles)
        self.regenerateGraph.clicked.connect(self.regenerateGraphFunc)
        self.OK.clicked.connect(self.outputGenerate)
        self.ProgressBarUpdater(0)
        self.activateSecondPol.stateChanged.connect(self.activateSecondPolFunc)
        
        self.activateSecondLorentzian.stateChanged.connect(self.activateSecondLorentzianFunc)
        self.activateFirstLorentzian.stateChanged.connect(self.activateFirstLorentzianFunc)

        self.plotWidget = pg.PlotWidget(Dialog)
        self.plotWidget.setGeometry(QtCore.QRect(10, 200, 1260, 550))
        self.activateSecondPolFunc()
        self.activateSecondLorentzianFunc()

    def activateFirstLorentzianFunc(self):

        if self.activateFirstLorentzian.isChecked() == True:
            self.label.setEnabled(False)
            self.bigPolyDegree.setEnabled(False)
        else:  
            self.label.setEnabled(True) 
            self.bigPolyDegree.setEnabled(True)

    def activateSecondLorentzianFunc(self):

        if self.activateSecondLorentzian.isChecked() == True:
            self.activateSecondPol.setCheckState(False)
            self.offsetOfPoly.setEnabled(True)
            self.label_2.setEnabled(True)
            self.windowsSizeLabel.setEnabled(True)
            self.windowsSize.setEnabled(True)
        else:
            self.offsetOfPoly.setEnabled(False)
            self.label_2.setEnabled(False)
            self.windowsSize.setEnabled(False)
            self.windowsSizeLabel.setEnabled(False)

    def activateSecondPolFunc(self):
        if self.activateSecondPol.isChecked() == True:
            #self.littlePolyOut.setEnabled(True)
            self.activateSecondLorentzian.setCheckState(False)
            self.offsetOfPoly.setEnabled(True)
            self.littlePolyDegree.setEnabled(True)
            self.label_2.setEnabled(True)
            self.label_3.setEnabled(True)
            self.windowsSizeLabel.setEnabled(True)
            self.windowsSize.setEnabled(True)
        else:
            self.littlePolyOut.setEnabled(False) 
            self.offsetOfPoly.setEnabled(False)
            self.littlePolyDegree.setEnabled(False)
            self.label_2.setEnabled(False)
            self.label_3.setEnabled(False)
            self.windowsSize.setEnabled(False)
            self.windowsSizeLabel.setEnabled(False)

    def showPlot(self,data):
        self.plotWidget.clear() 
        #self.plotWidget.plot(range(len(self.listOfPeaks)),self.listOfPeaks, pen='b', name='Noisy Data')
        self.plotWidget.plot(data.iloc[0,:], data.iloc[1,:], pen='w', name='PeakData') 

        # Set labels for the axes
        self.plotWidget.setLabel('left', 'Y Values')
        self.plotWidget.setLabel('bottom', 'X Values')
 
        # Grafiği çizmek için PyQtGraph PlotWidget kullanın
        self.plotWidget.setTitle('Plot')
        self.plotWidget.showGrid(True, True)

    fname =[]
    def browsefiles(self):
        fname=QFileDialog.getOpenFileName(None, 'Open file', os.getcwd(), 'Excel Files (*.xlsx)')
        self.filePathEdit.setText(fname[0])

    def regenerateGraphFunc(self):
        maxindex=self.find_nearest_index(self.listOfPeaks ,self.xmaxGraph.value())
        minindex=self.find_nearest_index(self.listOfPeaks ,self.xminGraph.value())
        dataList =self.listOfPeaks[maxindex:minindex]
        self.showPlot(pd.DataFrame(data=[range(len(dataList)),dataList]))

    def outputGenerate(self):
        self.OK.setEnabled(False)
        QtCore.QCoreApplication.processEvents()
        self.MakeGrafs()   

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Spectrum Analyzer"))
        self.label_2.setText(_translate("Dialog", "Offset (Pozitif olmalı)"))
        self.label_3.setText(_translate("Dialog", "Küçük Parçanın Polinom Derecesi"))
        self.label.setText(_translate("Dialog", "Genel Polinomun Derecesi"))
        self.OK.setText(_translate("Dialog", "Polinomu Üret"))
        self.fileLocate.setText(_translate("Dialog", "Dosya Belirt"))
        self.label_4.setText(_translate("Dialog", "X Maximum Değer"))
        self.label_5.setText(_translate("Dialog", "X Minimum Değer"))
        self.label_6.setText(_translate("Dialog", "Y Maximum Değer")) 
        self.label_7.setText(_translate("Dialog", "Y Minimum Değer"))
        
        self.windowsSizeLabel.setText(_translate("Dialog", "Ortalama alınacak deger sayısı"))
        self.regenerateGraph.setText(_translate("Dialog", "Yeniden Grafik yap"))
        self.bigPolyOut.setText(_translate("Dialog", "Genel Polinomu Çıktı Al"))
        self.littlePolyOut.setText(_translate("Dialog", "Parça Polinomu Çıktı Al"))
        self.peaksOut.setText(_translate("Dialog", "Tepe Değerleri Çıktı Al"))
        self.activateSecondPol.setText(_translate("Dialog", "2. polinomu aktif et"))
        
        self.activateSecondLorentzian.setText(_translate("Dialog", "2. Lorentzian Fit aktif et"))
        self.activateFirstLorentzian.setText(_translate("Dialog", "1. Lorentzian Fit aktif et"))

    def find_nearest_index(self,array, value):
        array = np.asarray(array)
        idx = (np.abs(array - value)).argmin()
        return idx 

    def get_index(self,array ,threshold, op): 
        if op == 1 :
            for index, value in enumerate(array):
                if value >= threshold: 
                    return index
        elif op == 0 :
            for index, value in reversed(list(enumerate(array))): 
                if value  <= threshold: 
                    return index
        return -1
     
    PolyDf = pd.DataFrame() 

    def ProgressBarUpdater(self,v):
        if v == 0 or  (self.progressBar.value() + v)>=100:
            if v>100:
                self.progressBar.setProperty("value",100)
            else:
                self.progressBar.setProperty("value",v)
        else: 
            self.progressBar.setProperty("value",self.progressBar.value() + v)
        QtCore.QCoreApplication.processEvents()
    def lorentzian(self,x, x0, gamma, A):
        return A * (gamma**2 / ((x - x0)**2 + gamma**2))

    def MakeGrafs(self):  

        tepeDegree = self.bigPolyDegree.value()
        degree = self.littlePolyDegree.value()
        offset = self.offsetOfPoly.value()
        file_path = self.filePathEdit.text() 
        
        print("EXCEL OKUMA BASLADI") 
        df = pd.read_excel(file_path)
        QtCore.QCoreApplication.processEvents()
        print("EXCEL OKUMA BITTI") 
        
        print("EXTINCTION HESAPlAMA BASLADI") 
        df=df.drop(df.index[0])
        df=df.drop(df.index[0])
        df=df.drop(df.index[1])
        df.reset_index(drop=True, inplace=True) 
        df.columns = df.iloc[0]
        df=df.drop(df.index[0])  
        df.reset_index(drop=True, inplace=True) 
        df.style.hide()
        df.columns = ['NM'] + [f'{col}{i+1}' for i, col in enumerate(df.columns[1:])]
        df= df.astype(float)
        df.iloc[:, 1:] = 1 - (10 ** (-df.iloc[:, 1:]))


        print("EXTINCTION HESAPlAMA BITTI")

        self.file_path = file_path[:-5]
        df.to_csv( self.file_path+"_extinction.csv",index=True,sep=';', decimal=',')

        self.PolyDf = df 

        print(df)
        time.sleep(1)
        self.ProgressBarUpdater(0)

        x = self.PolyDf.NM.to_numpy()
        self.nanometers = x
        columnsOfDf = self.PolyDf.columns[1:]
        x_copy = x.copy()
        self.ProgressBarUpdater(5)
        self.PolyDf=self.PolyDf.iloc[:,1:]
        self.listOfPeaks = [] 
        self.ProgressBarUpdater(0)
        A = 0.09463919098664426
        x0 = 274.3146806018545
        gamma = 585.8269681890237 

        initial_guess = [ A,x0, gamma]
        for c in self.PolyDf.columns:
        # Örnek veri oluşturma   
            self.ProgressBarUpdater(1)

            QtCore.QCoreApplication.processEvents()              
            x = x_copy
            y =self.PolyDf[c].to_numpy() 
            
            if self.activateFirstLorentzian.isChecked() == True:
                fit_params, _ = curve_fit(self.lorentzian, x, y, p0=initial_guess,maxfev=4000000)
                PolinomArray_temp = self.lorentzian(x, *fit_params)
            else:
                poly_temp = np.poly1d(np.polyfit(x, y, tepeDegree))
                PolinomArray_temp= poly_temp(x)

            maxIndex_temp = np.argmax(PolinomArray_temp) 
            self.PolyDf.loc[:,c] = PolinomArray_temp
            tepeX =  x[maxIndex_temp]
            # Polinomial Regresyon modelini uygulama
            if self.activateSecondPol.isChecked() == True or  self.activateSecondLorentzian.isChecked() == True:
                alt =  int(self.get_index(x,tepeX-offset,0))
                ust =  int(self.get_index(x,tepeX+offset,1))

                if(alt<0 or alt==-1):
                    alt=0
                if(ust>len(x) or  ust==-1):
                    ust= len(x)
                
                x=x[alt:ust]
                y=y[alt:ust]

                
                if self.activateSecondPol.isChecked() == True:
                    poly = np.poly1d(np.polyfit(x, y, degree))
                    PolinomArray= poly(x)
                    maxIndex = np.argmax(PolinomArray) 
                elif self.activateSecondLorentzian.isChecked() == True:
                    fit_params, _ = curve_fit(self.lorentzian, x, y, p0=initial_guess,maxfev=4000000)
                    y_fit = self.lorentzian(x, *fit_params)
                    maxIndex = np.argmax(y_fit)
                
                max_x =  x[maxIndex]
                max_y =  y_fit[maxIndex]

                print(max_x , max_y)

                self.listOfPeaks.append(max_x)  
            else:
                self.listOfPeaks.append(tepeX)

        if int(self.windowsSize.value()) > 0:
            self.listOfPeaks = self.apply_median_filter(self.listOfPeaks,int(self.windowsSize.value()))
        self.showPlot(pd.DataFrame(data=[range(len(self.listOfPeaks)),self.listOfPeaks]))

        self.ProgressBarUpdater(0)
        if(self.bigPolyOut.isChecked() == True):
            print("Big Poly Out Started")
            self.PolyDf = pd.concat([pd.Series(self.nanometers, name='NM'),  self.PolyDf], axis=1)
            self.PolyDf.transpose().to_csv(self.file_path+"_outputOfFirstPoly.csv",index=True,sep=';', decimal=',')
            print("Big Poly Out Finished")
            self.ProgressBarUpdater(50)
        if (self.peaksOut.isChecked() ==  True): 
            print("Peaks Out Started") 
            (pd.DataFrame(data=[columnsOfDf,self.listOfPeaks])).transpose().to_excel(self.file_path+"_outputOfPeakPoints.xlsx",index=False) 
            print("Peaks Out Finished")	
            self.ProgressBarUpdater(50)
        
        self.ProgressBarUpdater(100)
        time.sleep(1)  
        
        self.OK.setEnabled(True)
        print("Finished") 
        #exit()

    def apply_median_filter(self, y, window_size=3):
        filtered_y = []
        
        for i in range(len(y)):
            window_start = max(0, i - window_size)
            window_end = min(len(y), i + window_size + 1)
            median_value = np.median(y[window_start:window_end])
            filtered_y.append(median_value)
        
        return  filtered_y
   

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())

