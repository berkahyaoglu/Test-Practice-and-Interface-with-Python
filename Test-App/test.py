"""
This code was written by Berk Kahyaoğlu.
Email: berk.kahyaoglu99@gmail.com
In the code here, the vibration data in certain axes of an electrical machine are compared with each other and the difference is printed on the screen.
It can be optimized according to other data.
"""

from PyQt5.QtWidgets import *
from deneme1 import Ui_MainWindow
from PyQt5 import uic
import pandas as pd
import numpy as np
import warnings
import pyautogui
warnings.filterwarnings('ignore')

class main(QMainWindow):
    def __init__(self)-> None:
        super().__init__()
        uic.loadUi("deneme1.ui",self)
        self.qtTasarim = Ui_MainWindow()
        self.qtTasarim.setupUi(self)
        self.qtTasarim.pushButton_kontrol.clicked.connect(self.KontrolEt)
        self.qtTasarim.pushButton_yardim.clicked.connect(self.Yardim)
        self.qtTasarim.pushButton_kaydet.clicked.connect(self.Kaydet)
        self.qtTasarim.pushButton_gozat_2.clicked.connect(self.GozAt2)
        self.qtTasarim.pushButton_gozat.clicked.connect(self.GozAt)
        self.listele()
    def Kaydet(self):
        DosyaAdi2 = self.qtTasarim.lineEdit_dosya3.text()
        pyautogui.screenshot(f"{DosyaAdi2}.jpg",region=(490,150,950,490)).save(f"{DosyaAdi2}.jpg" ,"JPEG")
    def GozAt2(self):
        fname2 = QFileDialog.getOpenFileName(self,"Excel Dosyasını Seçin","C:Desktop\pro","xlsx files (*.xlsx)")
        self.qtTasarim.lineEdit_excel_2.setText(fname2[0]) 
    def GozAt(self):
        fname = QFileDialog.getOpenFileName(self,"Excel Dosyasını Seçin","C:Desktop\pro","xlsx files (*.xlsx)")
        self.qtTasarim.lineEdit_excel.setText(fname[0])
    def Yardim(self):
        QMessageBox.information(self,"İletişim","Volt Elektrik Motor Ar-Ge Merkezi \nMail Adresi: berk.kahyaoglu99@hotmail.com",QMessageBox.Close)
    def KontrolEt(self):
        ExcelAdi = self.qtTasarim.lineEdit_excel.text()
        ExcelAdi2 = self.qtTasarim.lineEdit_excel_2.text()
        a = pd.read_excel(f"{ExcelAdi2}").iloc[2:,0:3]
        b = pd.read_excel(f"{ExcelAdi2}").iloc[2:,6:8]
        c = pd.read_excel(f"{ExcelAdi2}").iloc[2:,51:53]
        d = pd.read_excel(f"{ExcelAdi2}").iloc[2:,56:58]
        saglikli_data = pd.concat([a,b,c,d],axis=1)
        saglikli_data.columns = ["Time","Z-Axis RMS Velocity (mm/sec)","Z-Axis Peak Velocity (mm/sec)","Z-Axis Crest Factor",
                                 "Z-Axis Kurtosis","X-Axis RMS Velocity (mm/sec)","X-Axis Peak Velocity (mm/sec)",
                                 "X-Axis Crest Factor","X-Axis Kurtosis"]
        zrms   = saglikli_data["Z-Axis RMS Velocity (mm/sec)"].mean()
        zpeak  = saglikli_data["Z-Axis Peak Velocity (mm/sec)"].mean()
        zcrest = saglikli_data["Z-Axis Crest Factor"].mean()
        zkurt  = saglikli_data["Z-Axis Kurtosis"].mean()
        xrms   = saglikli_data["X-Axis RMS Velocity (mm/sec)"].mean()
        xpeak  = saglikli_data["X-Axis Peak Velocity (mm/sec)"].mean()
        xcrest = saglikli_data["X-Axis Crest Factor"].mean()
        xkurt  = saglikli_data["X-Axis Kurtosis"].mean()
        veri = pd.DataFrame([zrms,zpeak,zcrest,zkurt,xrms,xpeak,xcrest,xkurt])
        self.qtTasarim.tableWidget.setItem(0,1,QTableWidgetItem(f"{round(zrms,2)}"))
        self.qtTasarim.tableWidget.setItem(1,1,QTableWidgetItem(f"{round(zpeak,2)}"))
        self.qtTasarim.tableWidget.setItem(2,1,QTableWidgetItem(f"{round(zcrest,2)}"))
        self.qtTasarim.tableWidget.setItem(3,1,QTableWidgetItem(f"{round(zkurt,2)}"))
        self.qtTasarim.tableWidget.setItem(4,1,QTableWidgetItem(f"{round(xrms,2)}"))
        self.qtTasarim.tableWidget.setItem(5,1,QTableWidgetItem(f"{round(xpeak,2)}"))
        self.qtTasarim.tableWidget.setItem(6,1,QTableWidgetItem(f"{round(xcrest,2)}"))
        self.qtTasarim.tableWidget.setItem(7,1,QTableWidgetItem(f"{round(xkurt,2)}"))
        
        a1 = pd.read_excel(f"{ExcelAdi}").iloc[2:,0:3]
        b1= pd.read_excel(f"{ExcelAdi}").iloc[2:,6:8]
        c1 = pd.read_excel(f"{ExcelAdi}").iloc[2:,51:53]
        d1 = pd.read_excel(f"{ExcelAdi}").iloc[2:,56:58]
        ölcülen_data = pd.concat([a1,b1,c1,d1],axis=1)
        ölcülen_data.columns = ["Time","Z-Axis RMS Velocity (mm/sec)","Z-Axis Peak Velocity (mm/sec)","Z-Axis Crest Factor",
                                 "Z-Axis Kurtosis","X-Axis RMS Velocity (mm/sec)","X-Axis Peak Velocity (mm/sec)",
                                 "X-Axis Crest Factor","X-Axis Kurtosis"]
        zrms1   = ölcülen_data["Z-Axis RMS Velocity (mm/sec)"].mean()
        zpeak1  = ölcülen_data["Z-Axis Peak Velocity (mm/sec)"].mean()
        zcrest1 = ölcülen_data["Z-Axis Crest Factor"].mean()
        zkurt1  = ölcülen_data["Z-Axis Kurtosis"].mean()
        xrms1   = ölcülen_data["X-Axis RMS Velocity (mm/sec)"].mean()
        xpeak1  = ölcülen_data["X-Axis Peak Velocity (mm/sec)"].mean()
        xcrest1 = ölcülen_data["X-Axis Crest Factor"].mean()
        xkurt1  = ölcülen_data["X-Axis Kurtosis"].mean()
        self.qtTasarim.tableWidget.setItem(0,2,QTableWidgetItem(f"{round(zrms1,2)}"))
        self.qtTasarim.tableWidget.setItem(1,2,QTableWidgetItem(f"{round(zpeak1,2)}"))
        self.qtTasarim.tableWidget.setItem(2,2,QTableWidgetItem(f"{round(zcrest1,2)}"))
        self.qtTasarim.tableWidget.setItem(3,2,QTableWidgetItem(f"{round(zkurt1,2)}"))
        self.qtTasarim.tableWidget.setItem(4,2,QTableWidgetItem(f"{round(xrms1,2)}"))
        self.qtTasarim.tableWidget.setItem(5,2,QTableWidgetItem(f"{round(xpeak1,2)}"))
        self.qtTasarim.tableWidget.setItem(6,2,QTableWidgetItem(f"{round(xcrest1,2)}"))
        self.qtTasarim.tableWidget.setItem(7,2,QTableWidgetItem(f"{round(xkurt1,2)}"))
    
        zrmsS   = (100*abs(zrms-zrms1))/zrms
        zpeakS  = (100*abs(zpeak-zpeak1))/zpeak
        zcrestS = (100*abs(zcrest-zcrest1))/zcrest
        zkurtS  = (100*abs(zkurt-zkurt1))/zkurt
        xrmsS   = (100*abs(xrms-xrms1))/xrms
        xpeakS  = (100*abs(xpeak-xpeak1))/xpeak
        xcrestS = (100*abs(xcrest-xcrest1))/xcrest
        xkurtS  = (100*abs(xkurt-xkurt1))/xkurt
        
        self.qtTasarim.tableWidget.setItem(0,3,QTableWidgetItem(f"%{round(zrmsS,2)}"))
        self.qtTasarim.tableWidget.setItem(1,3,QTableWidgetItem(f"%{round(zpeakS,2)}"))
        self.qtTasarim.tableWidget.setItem(2,3,QTableWidgetItem(f"%{round(zcrestS,2)}"))
        self.qtTasarim.tableWidget.setItem(3,3,QTableWidgetItem(f"%{round(zkurtS,2)}"))
        self.qtTasarim.tableWidget.setItem(4,3,QTableWidgetItem(f"%{round(xrmsS,2)}"))
        self.qtTasarim.tableWidget.setItem(5,3,QTableWidgetItem(f"%{round(xpeakS,2)}"))
        self.qtTasarim.tableWidget.setItem(6,3,QTableWidgetItem(f"%{round(xcrestS,2)}"))
        self.qtTasarim.tableWidget.setItem(7,3,QTableWidgetItem(f"%{round(xkurtS,2)}"))
        

        
    def listele(self):
        self.qtTasarim.tableWidget.setColumnWidth(0,300)
        self.qtTasarim.tableWidget.setColumnWidth(1,200)
        self.qtTasarim.tableWidget.setColumnWidth(2,200)
        self.qtTasarim.tableWidget.setColumnWidth(3,200)
        
        self.qtTasarim.tableWidget.setRowCount(8)
        
        self.qtTasarim.tableWidget.setItem(0,0,QTableWidgetItem("Z-Axis RMS Velocity (mm/sec)"))
        self.qtTasarim.tableWidget.setItem(1,0,QTableWidgetItem("Z-Axis Peak Velocity (mm/sec)"))
        self.qtTasarim.tableWidget.setItem(2,0,QTableWidgetItem("Z-Axis Crest Factor"))
        self.qtTasarim.tableWidget.setItem(3,0,QTableWidgetItem("Z-Axis Kurtosis"))
        self.qtTasarim.tableWidget.setItem(4,0,QTableWidgetItem("X-Axis RMS Velocity (mm/sec)"))
        self.qtTasarim.tableWidget.setItem(5,0,QTableWidgetItem("X-Axis Peak Velocity (mm/sec)"))
        self.qtTasarim.tableWidget.setItem(6,0,QTableWidgetItem("X-Axis Crest Factor"))
        self.qtTasarim.tableWidget.setItem(7,0,QTableWidgetItem("X-Axis Kurtosis"))
        
       
        
        
      
app = QApplication([])
pencere = main()
pencere.show()
app.exec_()


 

