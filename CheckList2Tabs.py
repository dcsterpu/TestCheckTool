from PyQt5.QtWidgets import *
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5 import QtWidgets
# from lxml import etree, objectify
import string
# import requests
import shutil
import sys
import zipfile
# import xlrd
import os
import openpyxl
from openpyxl import Workbook
import time
from datetime import date
# from openpyxl.styles import Alignment, Border, Side
# import win32com.client as win32
# import win32api


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
        Sheet = Workbook['Checks']

        self.list_document.append(Sheet['A1'].value)
        self.list_document.append(Sheet['A2'].value)
        self.list_document.append(Sheet['A3'].value)
        self.list_document.append(Sheet['A4'].value)
        self.list_document.append(Sheet['A5'].value)

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

    def buttonCheckClicked(self):
        for filename in self.list_document:
            exec("self.correspondences['" + filename + "'] = self.tab2.edit" + filename + ".text()")

        print(self.correspondences)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    app.exec_()
