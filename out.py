# -*- coding: utf-8 -*-

import matplotlib.pyplot as plt 
import pandas as pd
import numpy as np 

import sys
import time
from PyQt5.QtCore import QThread, QObject, pyqtSignal as Signal, pyqtSlot as Slot

import pyqtgraph as pg
from PyQt5 import QtCore, QtGui, QtWidgets 
from PyQt5.QtWidgets import QFileDialog
import os 

class Worker(QObject):
    completed = Signal(int)

    @Slot(str)
    def do_work(self, file_path):
        Ui_Dialog.PolyDf = pd.read_excel(file_path)
        self.completed.emit(100)

class Ui_Dialog(QObject):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(821, 590)

        self.graphicsView = QtWidgets.QGraphicsView(Dialog)
        self.graphicsView.setGeometry(QtCore.QRect(10, 260, 801, 281))
        self.graphicsView.setObjectName("graphicsView")

        self.gridLayoutWidget = QtWidgets.QWidget(Dialog)
        self.gridLayoutWidget.setGeometry(QtCore.QRect(10, 10, 431, 191))
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
        self.littlePolyDegree.setObjectName("littlePolyDegree")
        self.littlePolyDegree.setProperty("value", 6)
        self.gridLayout.addWidget(self.littlePolyDegree, 2, 1, 1, 1)
        self.bigPolyDegree = QtWidgets.QSpinBox(self.gridLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.bigPolyDegree.setFont(font)
        self.bigPolyDegree.setObjectName("bigPolyDegree") 
        self.bigPolyDegree.setProperty("value", 9)
        self.gridLayout.addWidget(self.bigPolyDegree, 0, 1, 1, 1)
    
        self.OK = QtWidgets.QPushButton(Dialog)
        self.OK.setGeometry(QtCore.QRect(590, 550, 221, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.OK.setFont(font)
        self.OK.setObjectName("OK")


        self.progressBar = QtWidgets.QProgressBar(Dialog)
        self.progressBar.setGeometry(QtCore.QRect(10, 550, 461, 23))
        self.progressBar.setProperty("value", 24)
        self.progressBar.setObjectName("progressBar")
        

        self.fileLocate = QtWidgets.QPushButton(Dialog)
        self.fileLocate.setGeometry(QtCore.QRect(460, 40, 351, 31))
        font = QtGui.QFont()
        font.setPointSize(12)
        self.fileLocate.setFont(font)
        self.fileLocate.setObjectName("fileLocate")


        self.filePathEdit = QtWidgets.QLineEdit(Dialog)
        self.filePathEdit.setGeometry(QtCore.QRect(460, 10, 351, 21))
        self.filePathEdit.setObjectName("filePathEdit")


        self.horizontalLayoutWidget = QtWidgets.QWidget(Dialog)
        self.horizontalLayoutWidget.setGeometry(QtCore.QRect(10, 220, 801, 41))
        self.horizontalLayoutWidget.setObjectName("horizontalLayoutWidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.horizontalLayoutWidget)
        self.horizontalLayout.setContentsMargins(0, 0, 0, 0)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_4 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_4.setFont(font)
        self.label_4.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_4.setObjectName("label_4")
        self.horizontalLayout.addWidget(self.label_4)

        self.maxGraph = QtWidgets.QDoubleSpinBox(self.horizontalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.maxGraph.setFont(font)
        self.maxGraph.setObjectName("maxGraph")
        self.horizontalLayout.addWidget(self.maxGraph)
        self.line = QtWidgets.QFrame(self.horizontalLayoutWidget)
        self.line.setFrameShape(QtWidgets.QFrame.VLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.horizontalLayout.addWidget(self.line)
        self.label_5 = QtWidgets.QLabel(self.horizontalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.label_5.setFont(font)
        self.label_5.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.label_5.setObjectName("label_5")
        self.horizontalLayout.addWidget(self.label_5)
        self.minGraph = QtWidgets.QDoubleSpinBox(self.horizontalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(14)
        self.minGraph.setFont(font)
        self.minGraph.setObjectName("minGraph")
        self.horizontalLayout.addWidget(self.minGraph)
        self.line_3 = QtWidgets.QFrame(self.horizontalLayoutWidget)
        self.line_3.setFrameShape(QtWidgets.QFrame.VLine)
        self.line_3.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_3.setObjectName("line_3")
        self.horizontalLayout.addWidget(self.line_3)
        self.regenerateGraph = QtWidgets.QPushButton(self.horizontalLayoutWidget)
        font = QtGui.QFont()
        font.setPointSize(12)
        self.regenerateGraph.setFont(font)
        self.regenerateGraph.setObjectName("regenerateGraph")
        self.horizontalLayout.addWidget(self.regenerateGraph)
        self.line_2 = QtWidgets.QFrame(Dialog)
        self.line_2.setGeometry(QtCore.QRect(10, 200, 801, 20))
        self.line_2.setFrameShape(QtWidgets.QFrame.HLine)
        self.line_2.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line_2.setObjectName("line_2")
        self.bigPolyOut = QtWidgets.QCheckBox(Dialog)
        self.bigPolyOut.setGeometry(QtCore.QRect(460, 80, 341, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.bigPolyOut.setFont(font)
        self.bigPolyOut.setObjectName("bigPolyOut")
        self.littlePolyOut = QtWidgets.QCheckBox(Dialog)
        self.littlePolyOut.setGeometry(QtCore.QRect(460, 110, 331, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.littlePolyOut.setFont(font)
        self.littlePolyOut.setObjectName("littlePolyOut")
        self.peaksOut = QtWidgets.QCheckBox(Dialog)
        self.peaksOut.setGeometry(QtCore.QRect(460, 140, 331, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.peaksOut.setFont(font)
        self.peaksOut.setObjectName("peaksOut")
        self.graphOut = QtWidgets.QCheckBox(Dialog)
        self.graphOut.setGeometry(QtCore.QRect(460, 170, 331, 31))
        font = QtGui.QFont()
        font.setPointSize(14)
        self.graphOut.setFont(font)
        self.graphOut.setObjectName("graphOut")
        self.graphOut.setChecked(True)
        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

        self.fileLocate.clicked.connect(self.browsefiles)
        self.regenerateGraph.clicked.connect(self.pngOut)
        self.OK.clicked.connect(self.outputGenerate)
        self.ProgressBarUpdater(0)
        self.plotWidget = pg.PlotWidget(Dialog)
        self.plotWidget.setGeometry(10, 260, 801, 281)
        
        self.regenerateGraph.clicked.connect(self.showPlot)  
    def showPlot(self):
        data = pd.DataFrame({'x': [1, 2, 3, 4, 5], 'y': [2, 4, 6, 8, 10]})
        
        # Grafiği çizmek için PyQtGraph PlotWidget kullanın
        self.plotWidget.plot(data['x'], data['y'], pen='b')
        self.plotWidget.setLabel('left', 'y')
        self.plotWidget.setLabel('bottom', 'x')
        self.plotWidget.setTitle('Plot')
        self.plotWidget.showGrid(True, True)

        
        # Display the plot in the QGraphicsView widget
        pixmap = QtGui.QPixmap('plot.png')
        scene = QtWidgets.QGraphicsScene(self.graphicsView)
        scene.addPixmap(pixmap)
        self.graphicsView.setScene(scene)

    fname =[]
    def browsefiles(self):
        fname=QFileDialog.getOpenFileName(None, 'Open file', os.getcwd(), 'Excel Files (*.xlsx)')
        self.filePathEdit.setText(fname[0])

    def pngOut(self):
        print("a")

    def outputGenerate(self):
        self.MakeGrafs()   

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "Dialog"))
        self.label_2.setText(_translate("Dialog", "Offset (Pozitif olmalı)"))
        self.label_3.setText(_translate("Dialog", "Küçük Parçanın Polinom Derecesi"))
        self.label.setText(_translate("Dialog", "Genel Polinomun Derecesi"))
        self.OK.setText(_translate("Dialog", "Polinomu Üret"))
        self.fileLocate.setText(_translate("Dialog", "Dosya Belirt"))
        self.label_4.setText(_translate("Dialog", "Maximum Değer"))
        self.label_5.setText(_translate("Dialog", "Minimum Değer"))
        self.regenerateGraph.setText(_translate("Dialog", "Yeniden Grafik Oluştur"))
        self.bigPolyOut.setText(_translate("Dialog", "Genel Polinomu Çıktı Al"))
        self.littlePolyOut.setText(_translate("Dialog", "Parça Polinomu Çıktı Al"))
        self.peaksOut.setText(_translate("Dialog", "Tepe Değerleri Çıktı Al"))
        self.graphOut.setText(_translate("Dialog", "Grafiği \".png\" Çıktı Al"))

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
    work_requested = Signal(str)
    isThreadFinished = False

    def on_finished(self,v):
        self.isThreadFinished=True
        self.ProgressBarUpdater(v)

    def ProgressBarUpdater(self,v):
        if v == 0 or  (self.progressBar.value() + v)>=100:
            if v>100:
                self.progressBar.setProperty("value",100)
            else:
                self.progressBar.setProperty("value",v)
        else: 
            self.progressBar.setProperty("value",self.progressBar.value() + v)
        QtCore.QCoreApplication.processEvents()

    def MakeGrafs(self):  

        tepeDegree = self.littlePolyDegree.value()
        degree = self.bigPolyDegree.value()
        offset = self.offsetOfPoly.value()
        file_path = self.filePathEdit.text()
        
        self.worker = Worker()
        self.worker_thread = QThread() 
        self.worker_thread.finished.connect(self.on_finished)
        self.worker.completed.connect(self.on_finished)
      
        self.work_requested.connect(self.worker.do_work)
        self.work_requested.emit(file_path) 

        self.worker.moveToThread(self.worker_thread) 
        self.worker_thread.start() 

        while self.isThreadFinished == False:
            self.ProgressBarUpdater(1)
            time.sleep(1)
            
        print("EXCEL OKUMA BITTI")
        self.isThreadFinished = False

        time.sleep(1)
        self.ProgressBarUpdater(0)
        column = len(self.PolyDf.iloc[0,:])
        row = len(self.PolyDf.iloc[:,0])
        
        x = self.PolyDf.NM.to_numpy()
        columnsOfDf = self.PolyDf.columns[1:]
        x_copy = x.copy()
        self.ProgressBarUpdater(5)
        self.PolyDf=self.PolyDf.iloc[:,1:]
        self.listOfPeaks = [] 
        self.ProgressBarUpdater(0)

        for c in self.PolyDf.columns:
        # Örnek veri oluşturma   
            self.ProgressBarUpdater(1)

            QtCore.QCoreApplication.processEvents()              
            x = x_copy
            y =self.PolyDf[c].to_numpy()
            #Grafik figürünü oluşturun
            poly_temp = np.poly1d(np.polyfit(x, y, tepeDegree))
            PolinomArray_temp= poly_temp(x)
            maxIndex_temp = np.argmax(PolinomArray_temp) 
            self.PolyDf[c] = PolinomArray_temp
            tepeX =  x[maxIndex_temp]
            # Polinomial Regresyon modelini uygulama
            
            alt =  int(self.get_index(x,tepeX-offset,0))
            ust =  int(self.get_index(x,tepeX+offset,1))

            if(alt<0 or alt==-1):
                alt=0
            if(ust>len(x) or  ust==-1):
                ust= len(x)
            
            x=x[alt:ust]
            y=y[alt:ust]

            poly = np.poly1d(np.polyfit(x, y, degree))
            PolinomArray= poly(x)
            maxIndex = np.argmax(PolinomArray) 
            max_x =  x[maxIndex]
            max_y =  PolinomArray[maxIndex]

            print(max_x , max_y)

            self.listOfPeaks.append(max_x) 
            
            
            # Grafik oluşturma   
            """
                plt.figure(figsize=(15,15))

                plt.plot(x,y, label=c ,color='orange',) 

                plt.plot(x, poly(x), color='blue', label=c)

                # Mutlak maksimumu işaretleme
                plt.plot (max_x,max_y, marker='o', markersize=8, color='green',label=('Mutlak Maksimum '+c))
                plt.axvline(x=max_x, color='green', linestyle='--')

                plt.xlabel('nm')
                plt.ylabel('abs')
                plt.legend()
                plt.show()
            """
         
        plt.figure(figsize=(200,25))
        plt.xticks( fontsize=10,rotation=90) 
        plt.plot(columnsOfDf, self.listOfPeaks)
        """
        #p = np.poly1d(np.polyfit(range(0,len( self.listOfPeaks)),  self.listOfPeaks,12))
        #plt.plot(columnsOfDf, p(range(0,len( self.listOfPeaks)))) 
        """
        self.ProgressBarUpdater(0)
        if(self.bigPolyOut.isChecked() == True):
            self.PolyDf.to_excel(os.getcwd()+"\\outputOfFirstPoly.xlsx",index=False)  
            self.ProgressBarUpdater(100)
        if (self.peaksOut.isChecked() ==  True):
            my_pd=pd.DataFrame(data=[columnsOfDf,self.listOfPeaks])
            my_pd.to_excel(os.getcwd()+"\\outputOfPeakPoints.xlsx",index=False) 
            self.ProgressBarUpdater(100)
        if (self.graphOut.isChecked() ==  True):
            plt.savefig(os.getcwd()+'\\tepeDegisim.png')
            self.ProgressBarUpdater(100)
        
        self.ProgressBarUpdater(100)
        time.sleep(1)  
        print("Finished")
        self.ProgressBarUpdater(0)
        #exit()
        

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Dialog = QtWidgets.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    sys.exit(app.exec_())
