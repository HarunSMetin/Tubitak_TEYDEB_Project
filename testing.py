
import matplotlib.pyplot as plt 
import pandas as pd
import numpy as np 
import openpyxl 

from scipy.optimize import curve_fit
def showPlot(data,listOfPeaks,title='Plot'):

    plt.figure()
    plt.title(title) 
    plt.plot(data.iloc[0,:], data.iloc[1,:], pen='w', name='PeakData')
    plt.plot(data.iloc[0,:], listOfPeaks, pen='b', name='Noisy Data')

    plt.xlabel('x')
    plt.ylabel('NM')
    plt.legend()
    plt.show() 

def get_index(array ,threshold, op): 
        if op == 1 :
            for index, value in enumerate(array):
                if value >= threshold: 
                    return index
        elif op == 0 :
            for index, value in reversed(list(enumerate(array))): 
                if value  <= threshold: 
                    return index
        return -1

def lorentzian(x, x0, gamma, A):
        return A * (gamma**2 / ((x - x0)**2 + gamma**2))  
  
def MakeGrafs(title,df,firstDegree,secondDegree,offset,file_path, activateSecondPol ,activateFirstLorentzian,activateSecondLorentzian, linspaceCount = 10000):  

        file_path = file_path[:-5]
        PolyDf = df.copy()
        print(PolyDf)
        x = PolyDf.NM.to_numpy()
        nanometers = x.copy()
        columnsOfDf = PolyDf.columns[1:] 
        PolyDf = PolyDf.iloc[:,1:]
        listOfPeaks = [] 

        A = 0.09463919098664426
        x0 = 274.3146806018545
        gamma = 585.8269681890237 

        initial_guess = [ A,x0, gamma]

        for c in PolyDf.columns: 
            x=nanometers
            y =PolyDf[c].to_numpy() 
            
            if activateFirstLorentzian:
                fit_params, _ = curve_fit(lorentzian, x, y, p0=initial_guess,maxfev=4000000)
                PolinomArray_temp = lorentzian(x, *fit_params)
            else:
                poly_temp = np.poly1d(np.polyfit(x, y, firstDegree))
                PolinomArray_temp= poly_temp(x)

            maxIndex_temp = np.argmax(PolinomArray_temp) 
            PolyDf.loc[:,c] = PolinomArray_temp
            tepeX =  x[maxIndex_temp]

            x_new1 = np.linspace(tepeX-2, tepeX+2, linspaceCount)

            if activateFirstLorentzian: 
                y_fit1 = lorentzian(x_new1, *fit_params)
            else:
                y_fit1= poly_temp(x_new1)
                    
            tepeX = x_new1[y_fit1.argmax()]  
            
            if activateSecondPol or activateSecondLorentzian :
                alt =  int(get_index(x,tepeX-offset,0))
                ust =  int(get_index(x,tepeX+offset,1))

                if(alt<0 or alt==-1):
                    alt=0
                if(ust>len(x) or  ust==-1):
                    ust= len(x)
                
                x=x[alt:ust]
                y=y[alt:ust]

                
                if activateSecondPol:
                    poly = np.poly1d(np.polyfit(x, y, secondDegree))
                    y_fit= poly(x) 
                    
                elif activateSecondLorentzian :
                    fit_params, _ = curve_fit( lorentzian, x, y, p0=initial_guess,maxfev=4000000)
                    y_fit = lorentzian(x, *fit_params)
               
                maxIndex = np.argmax(y_fit)
                max_x =  x[maxIndex]
                #max_y =  y_fit[maxIndex]
                
                x_new = np.linspace(max_x-2, max_x+2, linspaceCount)

                if activateSecondPol:
                   y_fit= poly(x_new)
                elif activateSecondLorentzian:
                    y_fit = lorentzian(x_new, *fit_params)
                max_x = x_new[y_fit.argmax()]
                #max_y = y_fit.max()
                print(max_x )

                listOfPeaks.append(max_x)  
            else:
                listOfPeaks.append(tepeX)
        
        showPlot(pd.DataFrame(data=[range(len(listOfPeaks)),listOfPeaks]),listOfPeaks,title)
        '''
        if(self.bigPolyOut.isChecked() == True):
            print("Big Poly Out Started")
            PolyDf = pd.concat([pd.Series( nanometers, name='NM'),  PolyDf], axis=1)
            PolyDf.transpose().to_csv(  file_path+"_outputOfFirstPoly.csv",index=True,sep=';', decimal=',')
            print("Big Poly Out Finished") 

        if (self.peaksOut.isChecked() ==  True): 
            print("Peaks Out Started") 
            (pd.DataFrame(data=[columnsOfDf, listOfPeaks])).transpose().to_excel( file_path+"_outputOfPeakPoints.xlsx",index=False) 
            print("Peaks Out Finished")	
        ''' 
        print("Finished") 
        #exit()


file_path="C:\\Users\\user\\Desktop\\test.xlsx" 
offset=10
activateSecondPol = True
activateFirstLorentzian = True
activateSecondLorentzian = True

df = pd.DataFrame()
print("EXCEL OKUMA BASLADI") 
df_all = pd.read_excel(file_path, sheet_name=None)
print("EXCEL OKUMA BITTI") 
sheet_names = df_all.sheet_names
for sheet_name in sheet_names:
    dftemp = df_all.parse(sheet_name)
    df = pd.concat([df, dftemp], axis=1)

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

print('Just 1. Lorentzian')
MakeGrafs('Just 1. Lorentzian',
          df,
          0,
          0,
          offset,
          file_path, 
          False,
          True,
          False) 

print('1. and 2. Lorentzian')
MakeGrafs('1. and 2. Lorentzian',
            df,
            0,
            0,
            offset,
            file_path,
            False,
            True,
            True)
    
for i in range(3,25): # first degree
    print('Just 1. Polynomial \n 1. degree: '+i)
    MakeGrafs('Just 1. Polynomial \n 1. degree: '+i,
                df,
                i,
                0,
                offset,
                file_path,
                False,
                False,
                False)
    for j in range(2,10): # second degree
        print('1. and 2. Polynomial \n 1. degree: '+i+' 2. degree: '+j)
        MakeGrafs('1. and 2. Polynomial \n 1. degree: '+i+' 2. degree: '+j,
                  df,
                i,
                j,
                offset,
                file_path,
                True ,
                False,
                False)
        
for i in range(2,10): # Lorentzian Polynomial
        print('1. Lorentzian and 2. Polynomial \n 2. degree: '+i )
        MakeGrafs('1. Lorentzian and 2. Polynomial \n 2. degree: '+i,
                  df,
                0,
                i,
                offset,
                file_path,
                True,
                True,
                False)

for i in range(3,25): # Polynomial Lorentzian
        print('1. Polynomial and 2. Lorentzian \n 1. degree: '+i )
        MakeGrafs('1. Polynomial and 2. Lorentzian \n 1. degree: '+i ,
                  df,
                i,
                0,
                offset,
                file_path,
                False,
                False,
                True)