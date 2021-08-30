import sys
from PyQt5 import QtWidgets 
from PyQt5.QtWidgets import *
from Ogrenciekle import *
from xlsxwriter.workbook import Workbook

Uygulama=QApplication(sys.argv)
penAna=QMainWindow()
ui=Ui_OgEkleme()
ui.setupUi(penAna)
penAna.show()

import sqlite3
global curs
global conn

conn=sqlite3.connect('veriler.db')
curs=conn.cursor()
sorguOgrenci=("CREATE TABLE IF NOT EXISTS obi(                 \
                Id INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT,     \
                Ogrno TEXT NOT NULL,                      \
                ad TEXT NOT NULL,                                  \
                soyad TEXT NOT NULL,                                \
                vizenot INTEGER NOT NULL,                \
                finalnot INTEGER NOT NULL,               \
                ort REAL NOT NULL,                     \
                harfnot TEXT NOT NULL)")
curs.execute(sorguOgrenci)
conn.commit()



def ekle():
    cevap=QMessageBox.question(penAna,"KAYIT EKLE","Kaydı eklemek istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        try:
            _ogno = ui.ogno.text()
            _ogad = ui.ogad.text()
            _ogsoyad = ui.ogsoyad.text()
            _ogvize = int(ui.ogvize.text())
            _ogfinal = int(ui.ogfinal.text())
            _ogort = (_ogvize * 0.4 + _ogfinal * 0.6)
            _ogharfnot =  ""

            if _ogort >= 85:
                _ogharfnot = "AA"
            elif _ogort >= 70:
                _ogharfnot = "BA"
            elif _ogort >= 60:
                _ogharfnot = "BB"
            elif _ogort >= 50:
                _ogharfnot = "CB"
            elif _ogort >= 40:
                _ogharfnot = "CC"
            else:
                _ogharfnot = "FF"
    
            curs.execute("INSERT INTO obi \
                     (Ogrno,ad,soyad,vizenot,finalnot,ort,harfnot) \
                      VALUES (?,?,?,?,?,?,?)", \
                      (_ogno, _ogad, _ogsoyad, _ogvize, _ogfinal, _ogort, _ogharfnot))
            conn.commit()
            liste()
            ui.statusbar.showMessage("Kayıt eklendi!",10000)
        except Exception as Hata:
            ui.statusbar.showMessage("Bir hata oluştu:"+str(Hata))
    else:
        ui.statusbar.showMessage("Ekleme işlemi iptal edildi.",10000)
   


def liste():
    
    ui.tableWidget.clear()
    
    ui.tableWidget.setHorizontalHeaderLabels(('Id', 'Öğrenci No', 'Ad', 'Soyad', 'Vize Notu', \
                                                'Final Notu', 'Ortalama', 'Harf Notu'))
    
    ui.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    
    curs.execute("SELECT * FROM obi")
    
    for satirIndeks, satirVeri in enumerate(curs):
        for sutunIndeks, sutunVeri in enumerate (satirVeri):
            ui.tableWidget.setItem(satirIndeks,sutunIndeks,QTableWidgetItem(str(sutunVeri)))
    ui.ogno.clear()
    ui.ogad.clear()
    ui.ogsoyad.clear()
    ui.ogvize.clear()
    ui.ogfinal.clear()
    
    curs.execute("SELECT COUNT(*) FROM obi")
    kayitSayisi=curs.fetchone()
    ui.label_3.setText(str(kayitSayisi[0]))

    curs.execute("SELECT AVG(ort) FROM obi")
    ortNot=curs.fetchone()
    ui.label_6.setText(str(ortNot[0]))

liste()

def Sil():
    cevap=QMessageBox.question(penAna,"KAYIT SİL","Kaydı silmek istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        secili=ui.tableWidget.selectedItems()
        silinecek=secili[1].text()
        try:
            curs.execute("DELETE FROM obi WHERE Ogrno='%s'" %(silinecek))
            conn.commit()
            
            liste()
            
            ui.statusbar.showMessage("Kayıt silindi!",10000)
        except Exception as Hata:
            ui.statusbar.showMessage("Bir hata oluştu:"+str(Hata))
    else:
        ui.statusbar.showMessage("Silme işlemi iptal edildi.",10000)

def excel():
    workbook = Workbook('obi.xlsx')
    worksheet = workbook.add_worksheet()
    isim_listesi=["ID", "Öğrenci No", "Ad", "Soyad", "Vize Notu", "Final Notu", "Ortalama", "Harf Notu"]
    for sutun,veri in enumerate(isim_listesi):
        worksheet.write(0,sutun,veri)
    c=conn.cursor()
    c.execute("select * from obi")
    mysel=c.execute("select * from obi")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet.write(i+1, j, row[j])
    workbook.close()

def Doldur():
    try:
        secili=ui.tableWidget.selectedItems()
        ui.ogno.setText(secili[1].text())
        ui.ogad.setText(secili[2].text())
        ui.ogsoyad.setText(secili[3].text())
        ui.ogvize.setText(secili[4].text())
        ui.ogfinal.setText(secili[5].text())
    except Exception as Hata:
            return

def Guncelle():
    cevap=QMessageBox.question(penAna,"KAYIT GÜNCELLEME","Kaydı güncellemek istediğinize emin misiniz?",\
                         QMessageBox.Yes | QMessageBox.No)
    if cevap==QMessageBox.Yes:
        try:
            secili=ui.tableWidget.selectedItems()
            _Id=int(secili[0].text())
            _ogno = ui.ogno.text()
            _ogad = ui.ogad.text()
            _ogsoyad = ui.ogsoyad.text()
            _ogvize = int(ui.ogvize.text())
            _ogfinal = int(ui.ogfinal.text())
            _ogort = (_ogvize * 0.4 + _ogfinal * 0.6)
            _ogharfnot =  ""
            if _ogort >= 85:
                _ogharfnot = "AA"
            elif _ogort >= 70:
                _ogharfnot = "BA"
            elif _ogort >= 60:
                _ogharfnot = "BB"
            elif _ogort >= 50:
                _ogharfnot = "CB"
            elif _ogort >= 40:
                _ogharfnot = "CC"
            else:
                _ogharfnot = "FF"

            
            curs.execute("UPDATE obi SET Ogrno=?,ad=?,soyad=?,vizenot=?,finalnot=?,ort=?,harfnot=? WHERE Id=?", \
                         (_ogno, _ogad, _ogsoyad, _ogvize, _ogfinal, _ogort, _ogharfnot, _Id))
            conn.commit()
            
            liste()
            
        except Exception as Hata:
            ui.statusbar.showMessage("Bir hata oluştu: "+str(Hata))
    else:
        ui.statusbar.showMessage("Güncelleme iptal edildi",10000)

ui.statusbar.showMessage("Öğrenci Takip Sistemi")
ui.oekle.clicked.connect(ekle)
ui.xlmolustur.clicked.connect(excel)
ui.osil.clicked.connect(Sil)
ui.guncelle.clicked.connect(Guncelle)
ui.tableWidget.itemSelectionChanged.connect(Doldur)


sys.exit(Uygulama.exec_())
