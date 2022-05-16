import sys
from PyQt5 import QtCore, QtGui, QtWidgets
from form import *
import time
from operator import mul,truediv
from petrovich.main import Petrovich
import openpyxl
from docxtpl import DocxTemplate

class MyWin(QtWidgets.QMainWindow):
    def __init__(self,parent=None):
        QtWidgets.QWidget.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.start()

        self.ui.leOrganization.textEdited.connect(self.Abb)
        self.ui.lePostGen.textEdited.connect(self.dublePostGen)
        self.ui.leSignetGen.textEdited.connect(self.dubleSignetGen)

        self.ui.leScoreAddres.textChanged.connect(self.AutoSet)
        self.ui.lePersonAccept.textChanged.connect(self.AutoSet)
        self.ui.lePersonAcceptPost.textChanged.connect(self.AutoSet)
        self.ui.leLegalAddres.textChanged.connect(self.AutoSet)

        self.ui.btn_Create.clicked.connect(self.create)

        self.ui.spinPriceCommon.textChanged.connect(self.changedPriceCommon)
        self.ui.btn_BasicPrice.clicked.connect(self.setBasicPrice)
        self.ui.btn_Price_Less.clicked.connect(self.priceCorrection)
        self.ui.btn_Price_More.clicked.connect(self.priceCorrection)

        self.ui.spinPrice051.textChanged.connect(self.price_to_m2)
        self.ui.spinPrice1275.textChanged.connect(self.price_to_m2)
        self.ui.spinPrice23.textChanged.connect(self.price_to_m2)
        self.ui.spinPrice46.textChanged.connect(self.price_to_m2)
        self.ui.spinPrice375.textChanged.connect(self.price_to_m2)
        self.ui.spinPrice45.textChanged.connect(self.price_to_m2)

        self.ui.spinNumber051.textChanged.connect(self.setBudget)
        self.ui.spinNumber1275.textChanged.connect(self.setBudget)
        self.ui.spinNumber23.textChanged.connect(self.setBudget)
        self.ui.spinNumber46.textChanged.connect(self.setBudget)
        self.ui.spinNumber375.textChanged.connect(self.setBudget)
        self.ui.spinNumber45.textChanged.connect(self.setBudget)
        self.ui.comboIter.currentTextChanged.connect(self.setBudget)


    def start(self):
        self.setTodayData()
        self.setBudget()
        self.installInfo()

    def installInfo(self):
        book = openpyxl.open("info.xlsx",read_only=True)
        sheet = book.active

        listProvider = []
        self.listPattern = []
        self.listPatternPath = []
        self.listPrice = []

        findProvider = False
        findPattern = False
        findPrice = False

        for i in range(1,sheet.max_row):
            if findProvider:
                if sheet[i+1][0].value == None:
                    findProvider = False
                listProvider.append(sheet[i][0].value)


            if findPattern:
                if sheet[i+1][0].value == None:
                    findPattern = False
                self.listPattern.append(sheet[i][0].value)
                self.listPatternPath.append(sheet[i][1].value)

            if findPrice:
                if sheet[i+1][0].value == None:
                    findPrice = False
                self.listPrice.append(sheet[i][0].value)


            if sheet[i][0].value == "Поставщики":
                    findProvider = True
            if sheet[i][0].value == "Договор":
                    findPattern = True
            if sheet[i][0].value == "Цены":
                    findPrice = True

        book.close()

        self.ui.comboProvider.addItems(listProvider)
        self.ui.comboContractPattern.addItems(self.listPattern)

    def price_to_m2(self,m:float):
        a = self.sender()

        if a == self.ui.spinPrice051:
            if not a.text():
                self.ui.label_m2_051.setText("0")
            else:
                self.ui.label_m2_051.setText("  " + str(round(float(str(a.value()).replace(",","."))/0.51,4)) + " р./м2")     # (str(  round(float(a.text()),4)) )
        elif a == self.ui.spinPrice1275:
            if not a.text():
                self.ui.label_m2_1275.setText("0")
            else:
                self.ui.label_m2_1275.setText("  " + str(round(float(str(a.value()).replace(",","."))/1.275,4)) + " р./м2")
        elif a == self.ui.spinPrice23:
            if not a.text():
                self.ui.label_m2_23.setText("0")
            else:
                self.ui.label_m2_23.setText("  " + str(round(float(str(a.value()).replace(",","."))/2.3,4)) + " р./м2")
        elif a == self.ui.spinPrice46:
            if not a.text():
                self.ui.label_m2_46.setText("0")
            else:
                self.ui.label_m2_46.setText("  " + str(round(float(str(a.value()).replace(",","."))/4.6,4)) + " р./м2")
        elif a == self.ui.spinPrice375:
            if not a.text():
                self.ui.label_m2_375.setText("0")
            else:
                self.ui.label_m2_375.setText("  " + str(round(float(str(a.value()).replace(",","."))/3.75,4)) + " р./м2")
        elif a == self.ui.spinPrice45:
            if not a.text():
                self.ui.label_m2_45.setText("0")
            else:
                self.ui.label_m2_45.setText("  " + str(round(float(str(a.value()).replace(",","."))/4.5,4)) + " р./м2")

        self.setBudget()

    def priceCorrection(self):
        action = self.sender()
        ratio = self.ui.spinPriceCorrection.value()
        ratio = ratio/100 + 1

        action_lib = {"+": mul, "-": truediv}
        all_price = {"self.ui.spinPrice051":self.ui.spinPrice051,
                     "self.ui.spinPrice1275":self.ui.spinPrice1275,
                     "self.ui.spinPrice23":self.ui.spinPrice23,
                     "self.ui.spinPrice46":self.ui.spinPrice46,
                     "self.ui.spinPrice375":self.ui.spinPrice375,
                     "self.ui.spinPrice45":self.ui.spinPrice45}

        for price in all_price.values():
            price.setValue(action_lib[action.text()](price.value(), ratio))

        self.setBudget()

    def setBasicPrice(self):
        self.ui.spinPrice051.setValue(int(self.listPrice[0]))
        self.ui.spinPrice1275.setValue(int(self.listPrice[1]))
        self.ui.spinPrice23.setValue(int(self.listPrice[2]))
        self.ui.spinPrice46.setValue(int(self.listPrice[3]))
        self.ui.spinPrice375.setValue(int(self.listPrice[4]))
        self.ui.spinPrice45.setValue(int(self.listPrice[5]))

        self.setBudget()

    def changedPriceCommon(self):
        m2 = self.ui.spinPriceCommon.value()

        self.ui.spinPrice051.setValue(0.51 * m2)
        self.ui.spinPrice1275.setValue(1.275 * m2)
        self.ui.spinPrice23.setValue(2.3 * m2)
        self.ui.spinPrice46.setValue(4.6 * m2)
        self.ui.spinPrice375.setValue(3.75 * m2)
        self.ui.spinPrice45.setValue(4.5 * m2)

        self.setBudget()

    def AutoSet(self):
        if self.ui.leScoreAddres.text() == " ":
            self.ui.leScoreAddres.setText(self.ui.leDeliveryAddres.text())

        if self.ui.lePersonAccept.text() == " ":
            self.ui.lePersonAccept.setText(self.ui.leSignet.text())

        if self.ui.lePersonAcceptPost.text() == " ":
            self.ui.lePersonAcceptPost.setText(self.ui.lePost.text())

        if self.ui.leLegalAddres.text() == " ":
            self.ui.leLegalAddres.setText(self.ui.leDeliveryAddres.text())

    def setBudget(self):
        iter_lib = {
            "1 раз в 2 недели"  :   2,
            "1 раз в неделю"    :   4,
            "2 раза в неделю"   :   8,
            "3 раза в неделю"   :   12,
            "ежедневно"         :   30
        }

        iter = iter_lib[self.ui.comboIter.currentText()]

        m1 = self.ui.spinNumber051.value() * self.ui.spinPrice051.value()
        m2 = self.ui.spinNumber1275.value() * self.ui.spinPrice1275.value()
        m3 = self.ui.spinNumber23.value() * self.ui.spinPrice23.value()
        m4 = self.ui.spinNumber46.value() * self.ui.spinPrice46.value()
        m5 = self.ui.spinNumber375.value() * self.ui.spinPrice375.value()
        m6 = self.ui.spinNumber45.value() * self.ui.spinPrice45.value()

        self.budget = (m1 + m2 + m3 + m4 + m5 + m6) * iter
        self.ui.label_8.setText(f"Сумма за 4 недели: {self.budget} руб.")

    def Abb(self):
        text = self.ui.leOrganization.text()

        if text.lower() == "ооо":
            self.ui.leOrganization.setText('Общество с ограниченной ответственностью "" ')
            self.ui.leOrganization.setText(self.ui.leOrganization.text()[:-1])
            self.ui.leOrganization.setCursorPosition(42)

        if text.lower() == "зао":
            self.ui.leOrganization.setText('Закрытое акционерное общество "" ')
            self.ui.leOrganization.setText(self.ui.leOrganization.text()[:-1])
            self.ui.leOrganization.setCursorPosition(31)

        if text.lower() == "ао":
            self.ui.leOrganization.setText('Акционерное общество "" ')
            self.ui.leOrganization.setText(self.ui.leOrganization.text()[:-1])
            self.ui.leOrganization.setCursorPosition(22)

        if text.lower() == "ип":
            self.ui.leOrganization.setText('Индивидуальный предприниматель ')
            self.ui.lePostGen.setText("индивидуального предпринимателя")

        short = self.ui.leOrganization.text()
        short = short.replace("Общество с ограниченной ответственностью","ООО")
        short = short.replace("Закрытое акционерное общество", "ЗАО")
        short = short.replace("Акционерное общество", "АО")
        short = short.replace("Индивидуальный предприниматель", "ИП")

        if "Общество с ограниченной ответственностью" in self.ui.leOrganization.text() or \
            "Закрытое акционерное общество" in self.ui.leOrganization.text() or \
            "Акционерное общество" in self.ui.leOrganization.text():
            self.ui.lePost.setText("Генеральный директор")
            self.ui.lePostGen.setText("генерального директора")
            self.ui.leFooting.setText("Устава")
        elif "Индивидуальный предприниматель" in self.ui.leOrganization.text():
            self.ui.lePost.setText("Индивидуальный предприниматель")
            self.ui.leFooting.setText("Свидетельства")

            signetip = self.ui.leOrganization.text()
            signetip = signetip.replace("Индивидуальный предприниматель ","")
            self.ui.leSignet.setText(signetip)


        self.ui.leOrgAbbreviated.setText(short)

    def dublePostGen(self):
        if self.ui.lePostGen.text() == " ":
            if "генеральный директор" in self.ui.lePost.text().lower():
                self.ui.lePostGen.setText("генерального директора")
            elif "индивидуальный предприниматель" in self.ui.leOrganization.text().lower():
                self.ui.lePostGen.setText("индивидуального предпринимателя")
            else:
                self.ui.lePostGen.setText(self.ui.lePost.text())

    def dubleSignetGen(self):
        if self.ui.leSignetGen.text() == " ":
            self.ui.leSignetGen.setText(self.ui.leSignet.text())

    def setTodayData(self):
        data = time.localtime()
        year = data[0]
        month = data[1]
        day = data[2]

        self.ui.dateContract.setDate(QtCore.QDate(year,month,day))

    def countMats(self):
        countMatLib =   {
            "mat1" : ("60x85 - ",self.ui.spinNumber051.value()),
            "mat2" : ("85x150 - ",self.ui.spinNumber1275.value()),
            "mat3" : ("115x200 - ",self.ui.spinNumber23.value()),
            "mat4" : ("115x400 - ",self.ui.spinNumber46.value()),
            "mat5" : ("150x250 - ",self.ui.spinNumber375.value()),
            "mat6" : ("150x300 - ",self.ui.spinNumber45.value())}

        self.allMats = ""

        for key,mats in countMatLib.items():
            if mats[1] > 0:
                self.allMats+= mats[0] + str(mats[1]) + " шт.\n"

    def create(self):
        self.countMats()

        self.all_names = {
        "number": self.ui.leContractNumber.text(),
        "data" :self.ui.dateContract.dateTime().toString("dd.MM.yyyy"),
        "organization": self.ui.leOrganization.text(),
        "abbCont": self.ui.leOrgAbbreviated.text(),
        "abbFile": self.ui.leOrgAbbreviated.text().replace('"',''),
        "signet": self.ui.leSignet.text(),
        "signetGen": self.ui.leSignetGen.text(),
        "post": self.ui.lePost.text(),
        "postGen": self.ui.lePostGen.text(),
        "footing": self.ui.leFooting.text(),
        "deliveryaddres": self.ui.leDeliveryAddres.text(),
        "orgTitle": self.ui.leOrgTitle.text(),
        "chart": self.ui.leChart.text(),
        "scoreAddres": self.ui.leScoreAddres.text(),
        "inn": self.ui.leInn.text(),
        "kpp": self.ui.leKpp.text(),
        "rscore": self.ui.leRscore.text(),
        "bank": self.ui.leBank.text(),
        "kscore": self.ui.leKscore.text(),
        "bik": self.ui.leBik.text(),
        "legalAddres": self.ui.leLegalAddres.text(),
        "accountant": self.ui.leAccountant.text(),
        "tel": self.ui.leTel.text(),
        "email": self.ui.leEmail.text(),
        "number1": self.ui.spinNumber051.text(),
        "number2": self.ui.spinNumber1275.text(),
        "number3" :self.ui.spinNumber23.text(),
        "number4": self.ui.spinNumber46.text(),
        "number5": self.ui.spinNumber375.text(),
        "number6": self.ui.spinNumber45.text(),
        "allMats": self.allMats,
        "color": str(self.ui.comboColor.currentText()[:-1] + "й"),
        "iter": self.ui.comboIter.currentText(),
        "personAccept": self.ui.lePersonAccept.text(),
        "personAcceptPost": self.ui.lePersonAcceptPost.text(),
        "personAcceptPostTel": self.ui.lePersonAcceptTel.text(),
        "commonTel": self.ui.leCommonTel.text(),
        "price1": self.ui.spinPrice051.text(),
        "price2": self.ui.spinPrice1275.text(),
        "price3": self.ui.spinPrice23.text(),
        "price4": self.ui.spinPrice46.text(),
        "price5": self.ui.spinPrice375.text(),
        "price6": self.ui.spinPrice45.text(),
        "provider": self.ui.comboProvider.currentText()}

        if self.ui.checkContract.isChecked():
            self.createContract()
        if self.ui.checkOrder.isChecked():
            self.createApplication()

    def createContract(self):
        patternIndex = self.ui.comboContractPattern.currentIndex()
        patternPath = "patterns\\" + self.listPatternPath[patternIndex]

        doc_cont = DocxTemplate(patternPath)
        doc_cont.render(self.all_names)
        doc_cont.save(f'contracts\\{self.all_names["number"]} - {self.all_names["abbFile"]}.docx')

    def createApplication(self):

        doc_ap = DocxTemplate('patterns\\ApplicationPattern.docx')
        doc_ap.render(self.all_names)
        doc_ap.save(f'applications\\{self.all_names["abbFile"]} - {self.all_names["orgTitle"]}.docx')




if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    myapp = MyWin()
    myapp.show()
    sys.exit(app.exec_())

