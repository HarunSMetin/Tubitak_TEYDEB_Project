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
from datetime import datetime


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
       
        self.line_2 = QtWidgets.QFrame(Dialog)
        self.line_2.setGeometry(QtCore.QRect(0, 145, 1280, 20))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")

        self.activateSecondPol = QtWidgets.QCheckBox(Dialog)
        self.activateSecondPol.setGeometry(QtCore.QRect(20, 125, 331, 25))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.activateSecondPol.setFont(font)
        self.activateSecondPol.setObjectName("activateSecondPol")

        self.activateSecondLorentzian = QtWidgets.QCheckBox(Dialog)
        self.activateSecondLorentzian.setGeometry(QtCore.QRect(250, 125, 331, 25))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.activateSecondLorentzian.setFont(font)
        self.activateSecondLorentzian.setObjectName("activateSecondLorentzian")

        self.activateFirstLorentzian = QtWidgets.QCheckBox(Dialog)
        self.activateFirstLorentzian.setGeometry(QtCore.QRect(480, 125, 331, 25))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.activateFirstLorentzian.setFont(font)
        self.activateFirstLorentzian.setObjectName("activateFirstLorentzian")

     
 
 

        self.labelLinspaceCount = QtWidgets.QLabel(Dialog)
        font.setPointSize(14)
        self.labelLinspaceCount.setFont(font)
        self.labelLinspaceCount.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.labelLinspaceCount.setObjectName("labelLinspaceCount")
        self.labelLinspaceCount.setGeometry(QtCore.QRect(400, 165, 180, 25)) 

        self.LinspaceCount = QtWidgets.QSpinBox(Dialog)
        self.LinspaceCount.setFont(font)
        self.LinspaceCount.setObjectName("LinspaceCount")
        self.LinspaceCount.setMaximum(1000000)
        self.LinspaceCount.setProperty("value", 20000)
        self.LinspaceCount.setGeometry(QtCore.QRect(550, 165, 110, 25)) 

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
            
            self.littlePolyOut.setEnabled(True)
            self.activateSecondPol.setCheckState(False)
            self.offsetOfPoly.setEnabled(True)
            self.label_2.setEnabled(True)
            self.windowsSizeLabel.setEnabled(True)
            self.windowsSize.setEnabled(True)
        else:
            
            self.littlePolyOut.setEnabled(False) 
            self.offsetOfPoly.setEnabled(False)
            self.label_2.setEnabled(False)
            self.windowsSize.setEnabled(False)
            self.windowsSizeLabel.setEnabled(False)

    def activateSecondPolFunc(self):
        if self.activateSecondPol.isChecked() == True:
            self.littlePolyOut.setEnabled(True)
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
        self.oldFilePath = self.filePathEdit.text() 
        fname=QFileDialog.getOpenFileName(None, 'Open file', os.getcwd(), 'Excel Files (*.xlsx)')
        self.filePathEdit.setText(fname[0]) 

 
    def outputGenerate(self):
        self.OK.setEnabled(False)
        QtCore.QCoreApplication.processEvents()
        self.MakeGrafs()   

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
        linspaceCount = self.LinspaceCount.value()

        tepeDegree = self.bigPolyDegree.value()
        degree = self.littlePolyDegree.value()
        offset = self.offsetOfPoly.value()
        file_path = self.filePathEdit.text() 
        if self.oldFilePath != file_path:
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
            df.iloc[:, 1:] = 1 - (1/(10 ** (df.iloc[:, 1:])))
            print("EXTINCTION HESAPlAMA BITTI")

            self.file_path = file_path[:-5]
            df.to_csv( self.file_path+"_extinction.csv",index=True,sep=';', decimal=',')
        else :  
            df=self.oldDf

        self.PolyDf = df 
        self.oldDf = df.copy()
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

        self.tepeX = -1
        alt_crop = 0
        ust_crop = len(x_copy)
        cropFlag = False

        for c in self.PolyDf.columns:
            
            y = self.PolyDf[c].to_numpy() 
            if self.tepeX != -1 :
                if cropFlag == False:
                    alt_crop =  int(self.get_index(x_copy,self.tepeX-60,0))
                    ust_crop =  int(self.get_index(x_copy,self.tepeX+60,1))

                    if(alt_crop<0 or alt_crop==-1):
                        alt_crop=0
                    if(ust_crop>len(x_copy) or  ust_crop==-1):
                        ust_crop= len(x_copy)

                    x_copy=x_copy[alt_crop:ust_crop] 
                    y= self.PolyDf[c].to_numpy() 
                    y=y[alt_crop:ust_crop] 

                    cropFlag = True
                else :  
                    y= self.PolyDf[c].to_numpy() 
                    y=y[alt_crop:ust_crop]      

            self.ProgressBarUpdater(1)
            QtCore.QCoreApplication.processEvents()              
            x = x_copy 
            if self.activateFirstLorentzian.isChecked() == True:
                fit_params, _ = curve_fit(self.lorentzian, x, y, p0=initial_guess,maxfev=4000000)
                PolinomArray_temp = self.lorentzian(x, *fit_params)
            else:
                poly_temp = np.poly1d(np.polyfit(x, y, tepeDegree))
                PolinomArray_temp= poly_temp(x)
           
            maxIndex_temp = np.argmax(PolinomArray_temp) 
            self.PolyDf.loc[alt_crop:ust_crop-1,c] = PolinomArray_temp
            self.tepeX =  x[maxIndex_temp]

            x_new1 = np.linspace(min(x),max(x), linspaceCount)
            if self.activateFirstLorentzian.isChecked() == True:
                y_fit1 = self.lorentzian(x_new1, *fit_params)
            else:
                y_fit1= poly_temp(x_new1)
                    
            self.tepeX = x_new1[y_fit1.argmax()]  
            
            if self.activateSecondPol.isChecked() == True or  self.activateSecondLorentzian.isChecked() == True:
                alt =  int(self.get_index(x_new1,self.tepeX-offset,0))
                ust =  int(self.get_index(x_new1,self.tepeX+offset,1))

                if(alt<0 or alt==-1):
                    alt=0
                if(ust>len(x_new1) or  ust==-1):
                    ust= len(x_new1)
                
                x=x_new1[alt:ust]
                y=y_fit1[alt:ust]

                
                if self.activateSecondPol.isChecked() == True:
                    poly = np.poly1d(np.polyfit(x, y, degree))
                    y_fit= poly(x) 
                    
                elif self.activateSecondLorentzian.isChecked() == True:
                    fit_params, _ = curve_fit(self.lorentzian, x, y, p0=initial_guess,maxfev=4000000)
                    y_fit = self.lorentzian(x, *fit_params)
               
                maxIndex = np.argmax(y_fit)
                max_x =  x[maxIndex]
                #max_y =  y_fit[maxIndex]
                
                x_new = np.linspace(max_x-offset, max_x+offset, linspaceCount)

                if self.activateSecondPol.isChecked() == True:
                   y_fit= poly(x_new)
                elif self.activateSecondLorentzian.isChecked() == True:
                    y_fit = self.lorentzian(x_new, *fit_params)


                xOfLittle = np.array(x_new)
                yOfLittle = np.array(y_fit)
                self.littlePolyOutDF = pd.DataFrame({'X': xOfLittle, 'Y': list(yOfLittle)}, columns=['X', 'Y'])

                max_x = x_new[y_fit.argmax()]
                max_y = y_fit.max()
                #print(max_x , max_y)
                if cropFlag:
                    self.listOfPeaks.append(max_x)  
            else:
                if cropFlag:
                    self.listOfPeaks.append(self.tepeX)

        if int(self.windowsSize.value()) > 0:
            self.listOfPeaks = self.apply_median_filter(self.listOfPeaks,int(self.windowsSize.value()))
        self.showPlot(pd.DataFrame(data=[range(len(self.listOfPeaks)),self.listOfPeaks]))

     
        out_filename=self.file_path     
        #add date time to filename
        now = str(datetime.datetime.now())    
        out_filename+=now.strftime("%Y_%m_%d_%H_%M_%S_%f")[:-3] 
        if self.activateFirstLorentzian.isChecked() == True:
            out_filename += "_firstLorentzian"
        else:
            out_filename += "_firstPol"+str(tepeDegree)
        if self.activateSecondPol.isChecked() == True :
            out_filename += "_secondPol"+str(degree)+"_degree"
        if self.activateSecondLorentzian.isChecked() == True:
            out_filename += "_secondLorentzian"
            
        
        self.ProgressBarUpdater(0)
        if(self.bigPolyOut.isChecked() == True):
            print("Big Poly Out Started")
            self.PolyDf = pd.concat([pd.Series(self.nanometers, name='NM'),  self.PolyDf], axis=1)
            self.PolyDf = self.PolyDf[alt_crop:ust_crop,:]
            self.PolyDf.transpose().to_csv(out_filename+"_outputOfFirstPoly.csv",index=True,sep=';', decimal=',')
            print("Big Poly Out Finished")
            self.ProgressBarUpdater(33)
        if(self.littlePolyOut.isChecked() == True):
            print("Little Poly Out Started")
            self.littlePolyOutDF.transpose().to_csv(out_filename+"_outputOfSecondPoly"+str(offset)+".csv",index=True,sep=';', decimal=',')
            print("Little Poly Out Finished")
            self.ProgressBarUpdater(33)

        if (self.peaksOut.isChecked() ==  True): 
            print("Peaks Out Started") 
            (pd.DataFrame(data=[columnsOfDf,self.listOfPeaks])).transpose().to_excel(self.file_path+"_outputOfPeakPoints.xlsx",index=False) 
            print("Peaks Out Finished")	
            self.ProgressBarUpdater(34)
        
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

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Spectrum Analyzer"))
        self.label_2.setText(_translate("Dialog", "Offset (Pozitif olmalı)"))
        self.label_3.setText(_translate("Dialog", "Küçük Parçanın Polinom Derecesi"))
        self.label.setText(_translate("Dialog", "Genel Polinomun Derecesi"))
        self.OK.setText(_translate("Dialog", "Polinomu Üret"))
        self.fileLocate.setText(_translate("Dialog", "Dosya Belirt"))
      
        
        self.windowsSizeLabel.setText(_translate("Dialog", "Ortalama alınacak deger sayısı"))
        
        self.bigPolyOut.setText(_translate("Dialog", "Genel Polinomu Çıktı Al"))
        self.littlePolyOut.setText(_translate("Dialog", "Parça Polinomu Çıktı Al"))
        self.peaksOut.setText(_translate("Dialog", "Tepe Değerleri Çıktı Al"))
        self.activateSecondPol.setText(_translate("Dialog", "2. polinomu aktif et"))
        
        self.activateSecondLorentzian.setText(_translate("Dialog", "2. Lorentzian Fit aktif et"))
        self.activateFirstLorentzian.setText(_translate("Dialog", "1. Lorentzian Fit aktif et")) 
        self.labelLinspaceCount.setText(_translate("Dialog", "Linspace Sayısı"))
   

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())

