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
# import win32api
import TestSource

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
        self.tabs.addTab(self.tab1, "Checklist")
        self.tabs.addTab(self.tab2, "Data files")
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
            if SheetSingleChecking.cell(index, 1).value is not None:
                current_list.append(SheetSingleChecking.cell(index, 1).value)
                current_list.append(SheetSingleChecking.cell(index, 2).value)
                if SheetSingleChecking.cell(index, 2).value == "No":
                    current_list.append("NT")
                else:
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

    def buttonCheckClicked(self):
        for filename in self.list_document:
            exec("self.correspondences['" + filename + "'] = self.tab2.edit" + filename + ".text().strip()")

        print(self.correspondences)

        for test in self.single_check_list:
            if test[1] == "Yes":
                if test[3] in self.dict_function:
                    nr_param = self.dict_function[test[3]]
                    if nr_param == 2 and test[4] is not None and test[5] is not None:
                        if test[3] == 'CheckDocumentTitle':
                            try:
                                if TestSource.CheckDocumentTitle(self, test[4], test[5]) is True:
                                    test[2] = 'OK'
                                else:
                                    test[2] = 'NOK'
                            except:
                                test[2] = 'NA'
                        elif test[3] == 'CheckHyperlink':
                            try:
                                if TestSource.CheckHyperlink(self, test[4], test[5]) is True:
                                    test[2] = 'OK'
                                else:
                                    test[2] = 'NOK'
                            except:
                                test[2] = 'NA'
                    elif nr_param == 3 and test[4] is not None and test[5] is not None and test[6] is not None:
                        if test[3] == 'CheckEqualValues':
                            try:
                                if TestSource.CheckEqualValues(self, test[4], test[5], test[6]) is True:
                                    test[2] = 'OK'
                                else:
                                    test[2] = 'NOK'
                            except:
                                test[2] = 'NA'
                        elif test[3] == 'CheckMultipleValues':
                            try:
                                if TestSource.CheckMultipleValues(self, test[4], test[5], test[6]) is True:
                                    test[2] = 'OK'
                                else:
                                    test[2] = 'NOK'
                            except:
                                test[2] = 'NA'
                    elif nr_param == 4 and test[4] is not None and test[5] is not None and test[6] is not None and test[7] is not None:
                        if test[3] == 'CheckDocInfoOrder':
                            try:
                                if TestSource.CheckDocInfoOrder(self, test[4], test[5], test[6], test[7]) is True:
                                    test[2] = 'OK'
                                else:
                                    test[2] = 'NOK'
                            except:
                                test[2] = 'NA'
                        elif test[3] == 'CheckNumberOfPoints':
                            try:
                                if TestSource.CheckNumberOfPoints(self, test[4], test[5], test[6], test[7]) is True:
                                    test[2] = 'OK'
                                else:
                                    test[2] = 'NOK'
                            except:
                                test[2] = 'NA'
                    elif nr_param == 5 and test[4] is not None and test[5] is not None and test[6] is not None and test[7] is not None and test[8] is not None:
                        if test[3] == 'CheckDocInfoParameter':
                            try:
                                if TestSource.CheckDocInfoParameter(self, test[4], test[5], test[6], test[7], test[8]) is True:
                                    test[2] = 'OK'
                                else:
                                    test[2] = 'NOK'
                            except:
                                test[2] = 'NA'
                    else:
                        test[2] = 'NA'
                else:
                    test[2] = 'NA'

        DocPath = self.tab1.edit1.text()
        Workbook = openpyxl.load_workbook(DocPath)
        SheetSingleChecking = Workbook['Single_Checking']

        row = 2
        for test in self.single_check_list:
            my_cell = SheetSingleChecking.cell(row, 3)
            my_cell.value = test[2]
            row += 1
        Workbook.save(DocPath)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    app.exec_()
