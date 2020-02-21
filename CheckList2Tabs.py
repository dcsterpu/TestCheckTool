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
        # self.tab1.setEnabled(False)


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

        tab.button = QPushButton("Start check", tab)
        tab.button.move(250, 400)
        tab.button.clicked.connect(self.buttonGenerateClicked)

        self.show()

    def openFileNameDialog1(self):
        fileName1, _filter = QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
        self.tab1.edit1.setText(fileName1)

    # def createFunctions(self):
    #     for function in self.list_document:
    #
    #         exec('''def openFileNameDialog''' + function + '''(self):
    #                 fileName''' + function + ''', _filter = QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
    #                 self.tab1.edit''' + function + '''.setText(''' + '\'' + '''fileName1'''+ function + '\'' + ''')''')


    # def initUIOptions(self, tab):
    #
    #     # set controls for Test Plan
    #     tab.label1 = QLabel("Test Plan", tab)
    #     tab.label1.move(20, 35)
    #     tab.edit1 = QLineEdit('', tab)
    #     tab.edit1.setDragEnabled(False)
    #     tab.edit1.move(120, 30)
    #     tab.edit1.resize(430, 25)
    #     tab.button1 = QPushButton("Browse", tab)
    #     tab.button1.move(560, 30)
    #     tab.button1.resize(50, 25)
    #     # tab.button1.clicked.connect(self.openFileNameDialog1)

    def createFunctions(self, data_list):


        # DocPath = self.tab1.edit1.text()
        # Workbook = openpyxl.load_workbook(DocPath)
        # Sheet = Workbook['Test']
        #
        # self.list_document.append(Sheet['A1'].value)
        # self.list_document.append(Sheet['A2'].value)
        # self.list_document.append(Sheet['A3'].value)
        # self.list_document.append(Sheet['A4'].value)
        # self.list_document.append(Sheet['A5'].value)


        for document_name in data_list:
            exec('''def openFileNameDialog''' + document_name + '''(self):
                                fileName''' + document_name + ''', _filter = QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
                                self.tab1.edit''' + document_name + '''.setText(''' + '\'' + '''fileName1''' + document_name + '\'' + ''')''')

    # exec('''def openFileNameDialog''' + 'Test' + '''(self):
    #                                 fileName''' + 'Test' + ''', _filter = QFileDialog.getOpenFileName(self, 'Open File', QtCore.QDir.rootPath(), '*.*')
    #                                 self.tab1.edit''' + 'Test' + '''.setText(''' + '\'' + '''fileName1''' + 'Test' + '\'' + ''')''')

    def buttonGenerateClicked(self):

        # self.createFunctions()

        DocPath = self.tab1.edit1.text()
        Workbook = openpyxl.load_workbook(DocPath)
        Sheet = Workbook['Test']

        self.list_document.append(Sheet['A1'].value)
        self.list_document.append(Sheet['A2'].value)
        self.list_document.append(Sheet['A3'].value)
        self.list_document.append(Sheet['A4'].value)
        self.list_document.append(Sheet['A5'].value)

        y_label = 35
        y_edit = 30
        self.createFunctions(self.list_document)

        for document_name in self.list_document:
            exec('self.tab2.label' + document_name + ' = QLabel(' + "\'" + document_name + "\'" + ', self.tab2)')
            exec('self.tab2.label' + document_name + '.move(20, ' + str(y_label) + ')')
            exec('self.tab2.edit' + document_name + ' = QLineEdit(' + "\'" + "\'" + ', self.tab2)')
            exec('self.tab2.edit' + document_name + '.setDragEnabled(False)')
            exec('self.tab2.edit' + document_name + '.move(120, ' + str(y_edit) + ')')
            exec('self.tab2.edit' + document_name + '.resize(430, 25)')
            exec('self.tab2.button' + document_name + ' = QPushButton("Browse", self.tab2)')
            exec('self.tab2.button' + document_name + '.move(560, ' + str(y_edit) + ')')
            exec('self.tab2.button' + document_name + '.resize(50, 25)')
            exec('self.tab2.button' + document_name + '.clicked.connect(self.openFileNameDialog' + document_name + ')')

            y_label += 30
            y_edit  += 30

        # self.createFunctions()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = Application()
    ex.show()
    app.exec_()
