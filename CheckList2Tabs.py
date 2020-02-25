from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import QtWidgets
from lxml import etree, objectify
import string
import requests
import shutil
import sys
import zipfile
# import xlrd
import os
import openpyxl
from openpyxl import Workbook
import time
from datetime import date
from openpyxl.styles import Alignment, Border, Side
import win32com.client as win32
import win32api


files_path = []
appName = "Check Tool"


class LineEdit(QLineEdit):

    def __init__(self, title, parent):
        super().__init__(title, parent)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [str(u.toLocalFile()) for u in event.mimeData().urls()]
        for f in files:
            temp = self.text()
            temp = temp + str(f) + "\n"
            files_path.append(str(f))
            self.setText(temp)
            # self.setText(("\n"))


class Application(QWidget):

    def __init__(self):
        super().__init__()

        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tabs.addTab(self.tab1, "Tab1")
        self.tabs.addTab(self.tab2, "Tab2")
        self.initUI(self.tab1)
        # self.initUIOptions(self.tab2)
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)
        self.setWindowTitle(appName)
        self.list_document = []
        self.correspondences = {}
        self.single_check_list = []
        self.dict_function = {
            "CheckEqualValues" : 3,
            "CheckDocumentTitle" : 2,
            "CheckDocInfoParameter" : 5,
            "CheckMultipleValues" : 3,
            "CheckHyperlink" : 2,
            "CheckDocInfoOrder" : 4,
            "CheckNumberOfPoints" : 4
        }

    def initUI(self, tab):

        # set controls for Checklist
        tab.label1 = QLabel("Checklist", tab)
        tab.label1.move(20, 35)
        tab.edit1 = QLineEdit('', tab)
        tab.edit1.setDragEnabled(False)
        tab.edit1.move(120, 30)
        tab.edit1.resize(430, 25)
        tab.button1 = QPushButton("Browse", tab)
        tab.button1.move(560, 30)
        tab.button1.resize(50, 25)
        tab.button1.clicked.connect(self.openFileNameDialog1)

        tab.button = QPushButton("Import Checklist", tab)
        tab.button.move(250, 400)
        tab.button.clicked.connect(self.buttonGenerateClicked)

        self.show()

    def openFileNameDialog1(self):
        fileName1, _filter = QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.edit1.setText(fileName1)

    def buttonGenerateClicked(self):

        DocPath = self.tab1.edit1.text()
        Workbook = openpyxl.load_workbook(DocPath)

        SheetSingleChecking = Workbook['Single_Checking']

        for index in range(2, SheetSingleChecking.max_row):
            current_list = []
            current_list.append(SheetSingleChecking.cell(index, 1).value)
            current_list.append(SheetSingleChecking.cell(index, 2).value)
            current_list.append(SheetSingleChecking.cell(index, 3).value)
            current_list.append(SheetSingleChecking.cell(index, 4).value)
            try:
                nr_param = self.dict_function[SheetSingleChecking.cell(index,4).value]
                if nr_param is not None:
                    for indexCol in range(5, 5 + nr_param):
                        current_list.append(SheetSingleChecking.cell(index, indexCol).value)
            except:
                pass
            self.single_check_list.append(current_list)

        Sheet = Workbook['Config']
        docColIndex = 0
        docRowIndex = 0

        for index1 in range(1, Sheet.max_row):
            for index2 in range(1, Sheet.max_column):
                if Sheet.cell(index1, index2).value == "Documents":
                    docRowIndex = index1
                    docColIndex = index2
                    break
            if docRowIndex != 0:
                break

        for index in range(docRowIndex + 1, Sheet.max_row):
            if Sheet.cell(index, docColIndex).value is not None:
                self.list_document.append(Sheet.cell(index, docColIndex).value)


        y_label = 35
        y_edit = 30
        for document_name in self.list_document:
            exec('self.tab2.label' + document_name + ' = QLabel(' + "\'" + document_name + "\'" + ', self.tab2)')
            exec('self.tab2.label' + document_name + '.move(20, ' + str(y_label) + ')')
            exec('self.tab2.edit' + document_name + ' = LineEdit(' + "\'" + "\'" + ', self.tab2)')
            exec('self.tab2.edit' + document_name + '.setDragEnabled(False)')
            exec('self.tab2.edit' + document_name + '.move(120, ' + str(y_edit) + ')')
            exec('self.tab2.edit' + document_name + '.resize(430, 25)')
            y_label += 30
            y_edit += 30

        self.tab2.button2 = QPushButton("Start Check", self.tab2)
        self.tab2.button2.move(250, 400)
        self.tab2.button2.clicked.connect(self.buttonCheckClicked)

    def download_file(self, url, user, password):

        try:
            os.stat(self.fileFolder)
        except:
            os.mkdir(self.fileFolder)
        try:
            response = requests.get(url, stream=True, auth=(user, password))
        except:
            return "Error"

        status = response.status_code
        FileName = response.headers['Content-Disposition'].split('"')[1]
        FilePath = self.fileFolder + FileName

        with open(FilePath, 'wb') as f:
            for chunk in response.iter_content(chunk_size=128):
                f.write(chunk)
        return FilePath

    def CheckEqualValues(self, Reference1, Reference2, Equal):

        if Reference1.split("<>")[0] in self.correspondences:
            DocPath1 = self.correspondences[Reference1.split("<>")[0]]
        else:
            DocPath1 = Reference1.split("<>")[0]

        if Reference2.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference2.split("<>")[0]]
        else:
            DocPath2 = Reference2.split("<>")[0]

        DocWorkbook1 = openpyxl.load_workbook(DocPath1)
        DocWorkbook2 = openpyxl.load_workbook(DocPath2)

        SheetName1 = Reference1.split("<>")[1]
        SheetName2 = Reference2.split("<>")[1]

        DocCel1 = Reference1.split("<>")[2]
        DocCel2 = Reference2.split("<>")[2]

        DocSheet1 = DocWorkbook1[SheetName1]
        CelValue1 = DocSheet1[DocCel1].value

        DocSheet2 = DocWorkbook2[SheetName2]
        CelValue2 = DocSheet2[DocCel2].value

        if Equal == "Yes":
            if CelValue1 == CelValue2:
                return True
            else:
                return False

        elif Equal == "No":
            if CelValue1 != CelValue2:
                return True
            else:
                return False

    def CheckDocumentTitle(self, Reference1, Reference2):

        if Reference1 in self.correspondences:
            Value1 = self.correspondences[Reference1].split("/")[-1]

        if Reference2.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference2.split("<>")[0]]
        else:
            DocPath2 = Reference2.split("<>")[0]

        DocWorkbook2 = openpyxl.load_workbook(DocPath2)
        DocSheet2 = DocWorkbook2[Reference2.split("<>")[1]]
        Value2 = DocSheet2[Reference2.split("<>")[2]].value

        if Value1.split(".")[0] == Value2.split(".")[0]:
            return True
        else:
            return False

    def CheckDocInfoParameter(self, Link, Parameter, Reference, User, Password):

        FilePath = self.download_file(Link, User, Password)

        extensions = ["xlsx", "xlsm"]
        if FilePath.split(".")[-1] in extensions:
            ext = FilePath.split(".")[0]
            with zipfile.ZipFile(FilePath, 'r') as zip_ref:
                zip_ref.extractall(ext)

            try:
                if os.path.isfile(ext + "\docProps\custom.xml"):
                    path = ext + "\docProps\custom.xml"
                    parser = etree.XMLParser(remove_comments=True)
                    tree = objectify.parse(path, parser=parser)
                    root = tree.getroot()
                    returned_parameter = root.find(".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = " + "\'" +Parameter + "\'" +"]/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
                    shutil.rmtree(ext, ignore_errors=True)
            except:
                shutil.rmtree(ext, ignore_errors=True)

        if Reference.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference.split("<>")[0]]
        else:
            DocPath2 = Reference.split("<>")[0]

        DocWorkbook2 = openpyxl.load_workbook(DocPath2)
        DocSheet2 = DocWorkbook2[Reference.split("<>")[1]]
        Value2 = DocSheet2[Reference.split("<>")[2]].value

        if returned_parameter in Value2:
            return True
        else:
            return False

    def CheckMultipleValues(self, List, Reference, Equal):

        list_values = List.split(";")

        if Reference.split("<>")[0] in self.correspondences:
            DocPath = self.correspondences[Reference.split("<>")[0]]
        else:
            DocPath = Reference.split("<>")[0]
        DocWorkbook = openpyxl.load_workbook(DocPath)
        DocSheet = DocWorkbook[Reference.split("<>")[1]]
        Value = DocSheet[Reference.split("<>")[2]].value

        if Equal == "True":
            if Value in list_values:
                return True
            else:
                return False
        elif Equal == "False":
            if Value not in list_values:
                return True
            else:
                return False

    def CheckHyperlink(self, Hyperlink, Reference):

        if Hyperlink.split("<>")[0] in self.correspondences:
            DocPath1 = self.correspondences[Hyperlink.split("<>")[0]]
        else:
            DocPath1 = Hyperlink.split("<>")[0]
        DocWorkbook1 = openpyxl.load_workbook(DocPath1, data_only= True, keep_vba=False)

        if Reference.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference.split("<>")[0]]
        else:
            DocPath2 = Reference.split("<>")[0]
        DocWorkbook2 = openpyxl.load_workbook(DocPath2, data_only= True, keep_vba=False)

        SheetName1 = Hyperlink.split("<>")[1]
        SheetName2 = Reference.split("<>")[1]

        DocCel1 = Hyperlink.split("<>")[2]
        DocCel2 = Reference.split("<>")[2]

        DocSheet1 = DocWorkbook1[SheetName1]
        CelValue1 = DocSheet1[DocCel1].hyperlink.target

        DocSheet2 = DocWorkbook2[SheetName2]
        CelValue2 = DocSheet2[DocCel2].value

        if CelValue1 == CelValue2:
            return True
        else:
            return False

    def CheckDocInfoOrder(self, DocInfoReference, Reference, User, Password):

        DocLinkIntranet = 'http://docinfogroupe.inetpsa.com/ead/doc/ref.' + DocInfoReference + '/v.vc/pj'
        DocLinkInternet = 'https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.' + DocInfoReference + '/v.vc/pj'


        FilePath = self.download_file(DocLinkIntranet, User, Password)
        if FilePath == "Error":
            FilePath = self.download_file(DocLinkInternet, User, Password)

        Doc1Name = FilePath.split("/")[-1]

        if Reference.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference.split("<>")[0]]
        else:
            DocPath2 = Reference.split("<>")[0]

        SheetName2 = Reference.split("<>")[1]
        DocCel2 = Reference.split("<>")[2]

        DocWorkbook2 = openpyxl.load_workbook(DocPath2, data_only=True, keep_vba=False)
        DocSheet2 = DocWorkbook2[SheetName2]
        CelValue2 = DocSheet2[DocCel2].value

        if Doc1Name == CelValue2:
            return True
        else:
            return False

    def CheckNumberOfPoints(self, Reference1, List, Reference2, Equal):

        if Reference1.split("<>")[0] in self.correspondences:
            DocPath1 = self.correspondences[Reference1.split("<>")[0]]
        else:
            DocPath1 = Reference1.split("<>")[0]

        if Reference2.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference2.split("<>")[0]]
        else:
            DocPath2 = Reference2.split("<>")[0]

        DocWorkbook1 = openpyxl.load_workbook(DocPath1)
        DocWorkbook2 = openpyxl.load_workbook(DocPath2)

        SheetName1 = Reference1.split("<>")[1]
        SheetName2 = Reference2.split("<>")[1]

        DocCel1 = Reference1.split("<>")[2]
        DocCel2 = Reference2.split("<>")[2]

        DocSheet1 = DocWorkbook1[SheetName1]
        CelValue1 = DocSheet1[DocCel1].value

        DocSheet2 = DocWorkbook2[SheetName2]
        CelValue2 = DocSheet2[DocCel2].value

        input_values = List.split(";")

        col = ""
        row = ""
        for char in DocCel1:
            if char.isalpha():
                col += char
            else:
                row += char

        row = int(row)
        number_points = 0

        if Equal == "Yes":
            while(DocSheet1[col + str(row)].value is not None):
                if DocSheet1[col + str(row)].value in input_values:
                    number_points += 1
                row += 1

        elif Equal == "No":
            while (DocSheet1[col + str(row)].value is not None):
                if DocSheet1[col + str(row)].value not in input_values:
                    number_points += 1
                row += 1

        if number_points == CelValue2:
            return True
        else:
            return False

    def buttonCheckClicked(self):
        for filename in self.list_document:
            exec("self.correspondences['" + filename + "'] = self.tab2.edit" + filename + ".text().strip()")

        print(self.correspondences)

        self.CheckEqualValues("C:\\Users\\msnecula\\Downloads\\TP_Checking_Tool_Request_200207.xlsx<>Header<>F1", "C:\\Users\\msnecula\\Downloads\\Checks.xlsx<>Checks<>B2", "No")

        self.CheckEqualValues("QIA_TP<>Header<>F1", "Checklist<>Checks<>B2", "No")

        self.CheckDocumentTitle("QIA_TP", "C:\\Users\\msnecula\\Downloads\\Checks.xlsx<>Checks<>A20")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    app.exec_()
