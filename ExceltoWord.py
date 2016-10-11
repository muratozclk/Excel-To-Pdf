# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created by: PyQt5 UI code generator 5.6
#
# WARNING! All changes made in this file will be lost!

import reportlab.rl_config
from reportlab.lib.enums import TA_RIGHT, TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.platypus import Image
from reportlab.platypus import KeepTogether
reportlab.rl_config.warnOnMissingFontGlyphs = 0
import pandas
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import QApplication
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QMainWindow
from reportlab.lib.units import inch, cm, mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import BaseDocTemplate, Frame, Paragraph, PageBreak, PageTemplate,SimpleDocTemplate
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import FrameBreak
from reportlab.platypus import NextPageTemplate
from reportlab.platypus import Spacer


class Example(QMainWindow):

    def __init__(self):
        super().__init__()
        self.initUI(self)



    def initUI(self,Form):

        Form.setObjectName("Form")
        Form.resize(550,350)
        self.pushButton = QtWidgets.QPushButton(Form)
        self.pushButton.setGeometry(QtCore.QRect(350, 250, 75, 25))
        self.pushButton.setObjectName("pushButton")
        self.pushButton_2 = QtWidgets.QPushButton(Form)
        self.pushButton_2.setGeometry(QtCore.QRect(450 , 250, 75, 25))
        self.pushButton_2.setObjectName("pushButton_2")
        self.listWidget = QtWidgets.QListWidget(Form)
        self.listWidget.setGeometry(QtCore.QRect(25, 30, 500, 175))
        self.listWidget.setObjectName("listWidget")
        self.lineEdit = QtWidgets.QLineEdit(Form)
        self.lineEdit.setGeometry(QtCore.QRect(25, 250, 75, 25))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit2 = QtWidgets.QLineEdit(Form)
        self.lineEdit2.setGeometry(QtCore.QRect(125, 250, 75, 25))
        self.lineEdit2.setObjectName("lineEdit")
        self.label = QtWidgets.QLabel(Form)
        self.label.setGeometry(QtCore.QRect(32, 225, 100, 25))
        self.label.setObjectName("label")
        self.label2 = QtWidgets.QLabel(Form)
        self.label2.setGeometry(QtCore.QRect(225, 10, 100, 25))
        self.label2.setObjectName("label")
        self.label3 = QtWidgets.QLabel(Form)
        self.label3.setGeometry(QtCore.QRect(130, 225, 100, 25))
        self.label3.setObjectName("label")



        self.retranslateUi(Form)
        QtCore.QMetaObject.connectSlotsByName(Form)


        self.pushButton.clicked.connect(self.dosyaEkle)
        self.pushButton_2.clicked.connect(self.cevirPDF)



    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Excel to Word"))
        self.pushButton.setText(_translate("Form", "Dosya Seç"))
        self.pushButton_2.setText(_translate("Form", "Çevir"))
        self.label.setText(_translate("Form", "Ders Sayısı"))
        self.label2.setText(_translate("Form", "Eklenen Dosyalar"))
        self.label3.setText(_translate("Form", "Boşluk Miktarı"))




    def dosyaEkle(self):

        self.dSayısı = int(self.lineEdit.text())
        self.bMiktari = self.lineEdit2.text()
        if self.bMiktari == "":
            self.bMiktari = 2
        self.bMiktari = int(self.bMiktari)

        self.filePath = ""
        self.filesPath = []
        for i in range(0,self.dSayısı):
            self.filePath, _  = QFileDialog.getOpenFileName(self, 'Dosya Seç', '') #Dosya yolunu seçtik
            self.listWidget.addItem(self.filePath)
            self.filesPath.append(self.filePath)


    def addPageNumber(self,canvas,doc):
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        pdfmetrics.registerFont(TTFont('ArialBI', 'ArialBI.ttf'))
        page_num = canvas.getPageNumber()
        text = "%s" % page_num
        canvas.saveState()
        canvas.setFont('ArialBd', 16)
        canvas.drawCentredString(10.6 * cm, 27.9* cm, self.kTürü)
        canvas.setFont('ArialBd',10)
        canvas.drawCentredString(10.6 * cm, 1.7 * cm,text)
        canvas.setFont('ArialBI',10)
        canvas.drawCentredString(17 * cm, 1 * cm, "Diğer Sayfaya Geçiniz.")
        canvas.line(10.6 * cm, 26 * cm, 10.6 * cm, 2.7 * cm)
        canvas.rotate(90)
        canvas.restoreState()


    def cevirPDF(self):
        for j in range(0,4):
            self.dersAdi = ""
            self.listSoru = []
            self.listA = []
            self.listB = []
            self.listC = []
            self.listD = []
            self.listE = []


            if j==0:
                self.kTürü='A'
            elif j==1:
                self.kTürü='B'
            elif j==2:
                self.kTürü='C'
            elif j==3:
                self.kTürü='D'
            Elements = []
            dAdi = 'Oturum-'+self.kTürü+'.pdf'
            doc = BaseDocTemplate(dAdi, showBoundary=0)

            for t in range(0,len(self.filesPath)):

                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=2, index_col=2, sheetname=j)
                for i in range(0, 25):
                    self.listSoru.append(self.gExcel.index[i])
                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=10, index_col=10, sheetname=j)
                for i in range(0, 25):
                    self.listA.append(self.gExcel.index[i])
                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=11, index_col=11, sheetname=j)
                for i in range(0, 25):
                    self.listB.append(self.gExcel.index[i])
                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=12, index_col=12, sheetname=j)
                for i in range(0, 25):
                    self.listC.append(self.gExcel.index[i])
                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=13, index_col=13, sheetname=j)
                for i in range(0, 25):
                    self.listD.append(self.gExcel.index[i])
                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=14, index_col=14, sheetname=j)
                for i in range(0, 25):
                    self.listE.append(self.gExcel.index[i])
                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=28, index_col=28, sheetname=j)
                self.dersAdi = self.gExcel.index[0]
                self.gExcel = pandas.read_excel(io=self.filesPath[t], parse_cols=27, index_col=27, sheetname=j)
                self.dersAdik = self.gExcel.index[0]
                for i in range(0, 25):
                    self.listSoru[i] = self.listSoru[i].replace('\n', '<br />\n')

                print("Excelden çekildi.")


                styles = getSampleStyleSheet()
                pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
                pdfmetrics.registerFont(TTFont('ArialBd', 'ArialBd.ttf'))

                # Baslık icin yazı ayarlarını yaptık
                titleStyle = styles["Title"]
                titleStyle.fontSize = 16
                titleStyle.fontName = 'ArialBd'
                titleStyle.leading = 65

                #Soru için font
                arStyle = styles["Heading1"]
                arStyle.fontSize = 10
                arStyle.fontName = 'ArialBd'
                arStyle.leading = 11
                arStyle.alignment = TA_JUSTIFY


                # Paragraflar için yazı tipini ayarladık
                parStyle = styles["Normal"]
                parStyle.fontSize = 10
                parStyle.fontName = 'Arial'
                parStyle.leading = 16
                #parStyle.alignment = TA_LEFT

                # Acıklama için yazı tipini ayarladık
                parStyle2 = styles["Normal"]
                parStyle2.fontSize = 10
                parStyle2.fontName = 'Arial'
                parStyle2.leading = 16
                parStyle2.alignment = TA_JUSTIFY

                # column için ölçüleri oluşturduk
                frameHeight = doc.height + 2 * inch
                firstPageHeight = 6 * inch
                firstPageBottom = frameHeight - firstPageHeight

                # Baslık frame'i
                frameT = Frame(2.7 * cm, firstPageBottom, doc.width, firstPageHeight)

                # Two column - burda oluşturduk
                frame1 = Frame(2.7 * cm,2.7 * cm,7.5 * cm,23.3 * cm, id='col1')
                frame2 = Frame(11.6 * cm,2.7 * cm,7.5 * cm,23.3 * cm, id='col2')

                Elements.append(Paragraph('<br />\n'+self.dersAdi, titleStyle))
                Elements.append(NextPageTemplate('TwoCol'))
                Elements.append(FrameBreak())

                for i in range(0, 25):
                    if i != 24:
                        index = self.listSoru[i].rfind("\n")
                        imagel=self.listSoru[i].find("&")
                        image2 = self.listSoru[i].rfind("&")
                        if imagel == -1 or image2 == -1:
                            if index == -1:
                                ptext = '<font fontSize=10 fontName=ArialBd>%s.&nbsp;&nbsp;</font>' % str(i+1)
                                pptext = ptext+self.listSoru[i]
                                Soru = Paragraph(pptext, arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                Elements.append(KeepTogether([Soru, Bosluk, A, B, C, D, E, Bosluk2]))
                            else:
                                soru1 = self.listSoru[i][0:index]
                                soru2 = self.listSoru[i][index:]
                                ptext = '<font fontSize=10 fontName=ArialBd>%s.&nbsp;&nbsp;</font>' % str(i + 1)
                                pptext = ptext + soru1
                                Soruy1 = Paragraph(pptext, parStyle2)
                                Soruy2 = Paragraph(soru2,arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                Elements.append(KeepTogether([Soruy1,Soruy2, Bosluk, A, B, C, D, E, Bosluk2]))
                        else:
                            imageAd=self.listSoru[i][imagel+1:image2]
                            imageK=self.listSoru[i][imagel:image2+1]
                            soru1 = self.listSoru[i][0:index]
                            soruBas = soru1[0:imagel]
                            soruSon = soru1[image2+1:]
                            soru2 = self.listSoru[i][index:]
                            imagePath="C:\Image\\" + imageAd
                            im = Image(imagePath,7.5 * cm, 5*cm)
                            ySoru = self.listSoru[i]
                            yySoru = ySoru.replace(imageK,"")
                            if index == -1:
                                ptext = Paragraph(str(i+1)+".&nbsp;&nbsp;",arStyle)
                                Soru = Paragraph(yySoru, arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                Elements.append(KeepTogether([ptext,im,Soru, Bosluk, A, B, C, D, E, Bosluk2]))
                            else:

                                ptext = '<font fontSize=10 fontName=ArialBd>%s.&nbsp;&nbsp;</font>' % str(i + 1)
                                pptext = ptext + soruBas
                                soruy1Bas = Paragraph(pptext, parStyle2)
                                soruy1Son = Paragraph(soruSon,parStyle2)
                                Soruy2 = Paragraph(soru2,arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                Elements.append(KeepTogether([soruy1Bas,im,soruy1Son,Soruy2, Bosluk, A, B, C, D, E, Bosluk2]))
                    else:
                        index = self.listSoru[i].rfind("\n")
                        imagel = self.listSoru[i].find("&")
                        image2 = self.listSoru[i].rfind("&")
                        if imagel == -1 or image2 == -1:
                            if index == -1:
                                ptext = '<font fontSize=10 fontName=ArialBd>%s.&nbsp;&nbsp;</font>' % str(i + 1)
                                pptext = ptext + self.listSoru[i]
                                Soru = Paragraph(pptext, arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                tBitti = Paragraph(self.dersAdi + ' TESTİ BİTTİ.', arStyle)
                                Elements.append(KeepTogether([Soru, Bosluk, A, B, C, D, E, Bosluk2,tBitti]))

                            else:
                                soru1 = self.listSoru[i][0:index]
                                soru2 = self.listSoru[i][index:]
                                ptext = '<font fontSize=10 fontName=ArialBd>%s.&nbsp;&nbsp;</font>' % str(i + 1)
                                pptext = ptext + soru1
                                Soruy1 = Paragraph(pptext, parStyle2)
                                Soruy2 = Paragraph(soru2, arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                tBitti = Paragraph(self.dersAdi + ' TESTİ BİTTİ.', arStyle)
                                Elements.append(KeepTogether([Soruy1, Soruy2, Bosluk, A, B, C, D, E, Bosluk2,tBitti]))

                        else:
                            imageAd = self.listSoru[i][imagel + 1:image2]
                            imageK = self.listSoru[i][imagel:image2 + 1]
                            soru1 = self.listSoru[i][0:index]
                            soruBas = soru1[0:imagel]
                            soruSon = soru1[image2 + 1:]
                            soru2 = self.listSoru[i][index:]
                            imagePath = "D:\İçerikler\Python\ExcelToWord\Image\\" + imageAd
                            im = Image(imagePath, 7.5 * cm, 5 * cm)
                            ySoru = self.listSoru[i]
                            yySoru = ySoru.replace(imageK, "")
                            if index == -1:
                                ptext = Paragraph(str(i + 1) + ".&nbsp;&nbsp;", arStyle)
                                Soru = Paragraph(yySoru, arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                tBitti = Paragraph(self.dersAdi + ' TESTİ BİTTİ.', arStyle)
                                Elements.append(KeepTogether([ptext, im, Soru, Bosluk, A, B, C, D, E, Bosluk2,tBitti]))

                            else:

                                ptext = '<font fontSize=10 fontName=ArialBd>%s.&nbsp;&nbsp;</font>' % str(i + 1)
                                pptext = ptext + soruBas
                                soruy1Bas = Paragraph(pptext, parStyle2)
                                soruy1Son = Paragraph(soruSon, parStyle2)
                                Soruy2 = Paragraph(soru2, arStyle)
                                Bosluk = Spacer(1, 0.3 * cm)
                                A = Paragraph('A)&nbsp;&nbsp;' + str(self.listA[i]), parStyle)
                                B = Paragraph('B)&nbsp;&nbsp;' + str(self.listB[i]), parStyle)
                                C = Paragraph('C)&nbsp;&nbsp;' + str(self.listC[i]), parStyle)
                                D = Paragraph('D)&nbsp;&nbsp;' + str(self.listD[i]), parStyle)
                                E = Paragraph('E)&nbsp;&nbsp;' + str(self.listE[i]), parStyle)
                                Bosluk2 = Spacer(1, self.bMiktari * cm)
                                tBitti = Paragraph(self.dersAdi+' TESTİ BİTTİ.',arStyle)
                                Elements.append(KeepTogether([soruy1Bas, im, soruy1Son, Soruy2, Bosluk, A, B, C, D, E, Bosluk2,tBitti]))


                Elements.append(NextPageTemplate('Title'))
                Elements.append(PageBreak())

                doc.addPageTemplates([PageTemplate(id='Title', frames=[frameT, frame1, frame2], onPage=self.addPageNumber),
                                      PageTemplate(id='TwoCol', frames=[frame1, frame2], onPage=self.addPageNumber), ])

                self.listSoru.clear()
                self.listA.clear()
                self.listB.clear()
                self.listC.clear()
                self.listD.clear()
                self.listE.clear()

            doc.build(Elements)

            print("PDF'e çevrildi.")



if __name__ == '__main__':
    import sys

    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Example()
    ui.initUI(Form)
    Form.show()
    sys.exit(app.exec_())