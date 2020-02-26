import os
import shutil
import zipfile
import openpyxl
from lxml import etree, objectify

import CheckList2Tabs


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

    if Equal is True:
        if CelValue1 == CelValue2:
            return True
        else:
            return False

    elif Equal is False:
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
                returned_parameter = root.find(
                    ".//{http://schemas.openxmlformats.org/officeDocument/2006/custom-properties}property[@name = " + "\'" + Parameter + "\'" + "]/{http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes}lpwstr").text
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

    if Equal.casefold() == "true":
        if Value in list_values:
            return True
        else:
            return False
    elif Equal.casefold() == "false":
        if Value not in list_values:
            return True
        else:
            return False


def CheckHyperlink(self, Hyperlink, Reference):
    if Hyperlink.split("<>")[0] in self.correspondences:
        DocPath1 = self.correspondences[Hyperlink.split("<>")[0]]
    else:
        DocPath1 = Hyperlink.split("<>")[0]
    DocWorkbook1 = openpyxl.load_workbook(DocPath1, data_only=True, keep_vba=False)

    if Reference.split("<>")[0] in self.correspondences:
        DocPath2 = self.correspondences[Reference.split("<>")[0]]
    else:
        DocPath2 = Reference.split("<>")[0]
    DocWorkbook2 = openpyxl.load_workbook(DocPath2, data_only=True, keep_vba=False)

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

    if Equal.casefold() == "true":
        while (DocSheet1[col + str(row)].value is not None):
            if DocSheet1[col + str(row)].value in input_values:
                number_points += 1
            row += 1

    elif Equal.casefold() == "false":
        while (DocSheet1[col + str(row)].value is not None):
            if DocSheet1[col + str(row)].value not in input_values:
                number_points += 1
            row += 1

    if number_points == CelValue2:
        return True
    else:
        return False