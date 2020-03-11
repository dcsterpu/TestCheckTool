import os
import shutil
import zipfile
import openpyxl
from lxml import etree, objectify
import CheckList2Tabs


def CheckEqualValues(self, Reference1, Reference2, Equal):

    try:
        if Reference1.split("<>")[0] in self.correspondences:
            DocPath1 = self.correspondences[Reference1.split("<>")[0]]
        else:
            DocPath1 = Reference1.split("<>")[0]

        if Reference2.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference2.split("<>")[0]]
        else:
            DocPath2 = Reference2.split("<>")[0]

        DocWorkbook1 = openpyxl.load_workbook(DocPath1, data_only=True)
        DocWorkbook2 = openpyxl.load_workbook(DocPath2, data_only=True)

        SheetName1 = Reference1.split("<>")[1]
        SheetName2 = Reference2.split("<>")[1]

        DocCel1 = Reference1.split("<>")[2]
        DocCel2 = Reference2.split("<>")[2]

        DocSheet1 = DocWorkbook1[SheetName1]
        CelValue1 = DocSheet1[DocCel1].value

        DocSheet2 = DocWorkbook2[SheetName2]
        CelValue2 = DocSheet2[DocCel2].value


        if str(Equal).casefold() == "true":
            if CelValue1 == CelValue2:
                return 'OK'
            else:
                return 'NOK'

        elif str(Equal).casefold() == "false":
            if CelValue1 != CelValue2:
                return 'OK'
            else:
                return 'NOK'
    except:
        return 'NA'

def CheckDocumentTitle(self, Reference1, Reference2):
    try:
        if Reference1 in self.correspondences:
            Value1 = self.correspondences[Reference1].split("/")[-1]

        if Reference2.split("<>")[0] in self.correspondences:
            DocPath2 = self.correspondences[Reference2.split("<>")[0]]
        else:
            DocPath2 = Reference2.split("<>")[0]

        DocWorkbook2 = openpyxl.load_workbook(DocPath2, data_only=True)
        DocSheet2 = DocWorkbook2[Reference2.split("<>")[1]]
        Value2 = DocSheet2[Reference2.split("<>")[2]].value

        if Value1.split(".")[0] == Value2.split(".")[0]:
            return 'OK'
        else:
            return 'NOK'
    except:
        return 'NA'

def CheckDocInfoParameter(self, DocInfoReference, NumberFile, Parameter):
    try:
        if NumberFile is None:
            DocLinkIntranet = 'http://docinfogroupe.inetpsa.com/ead/doc/ref.' + DocInfoReference + '/v.vc/pj'
            DocLinkInternet = 'https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.' + DocInfoReference + '/v.vc/pj'
        else:
            DocLinkIntranet = 'http://docinfogroupe.inetpsa.com/ead/doc/ref.' + DocInfoReference + '/v.vc/nPj.' + str(NumberFile)
            DocLinkInternet = 'https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.' + DocInfoReference + '/v.vc/nPj.' + str(NumberFile)


        User = self.tab1.TextBoxUser.text()
        Password = self.tab1.TextBoxPass.text()
        FilePath = self.download_file(DocLinkIntranet, User, Password)
        if FilePath == "Error":
            FilePath = self.download_file(DocLinkInternet, User, Password)

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

        return returned_parameter
    except:
        return 'NA'

def CheckHyperlink(self, Hyperlink, Reference):
    try:
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
            return 'OK'
        else:
            return 'NOK'
    except:
        return 'NA'

def CheckDocInfoOrder(self, DocInfoReference, Reference):
    try:
        DocLinkIntranet = 'http://docinfogroupe.inetpsa.com/ead/doc/ref.' + DocInfoReference + '/v.vc/pj'
        DocLinkInternet = 'https://docinfogroupe.psa-peugeot-citroen.com/ead/doc/ref.' + DocInfoReference + '/v.vc/pj'

        User = self.tab1.TextBoxUser.text()
        Password = self.tab1.TextBoxPass.text()
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
            return 'OK'
        else:
            return 'NOK'
    except:
        return 'NA'

def CountNumberOfPoints(self, Reference, Column, Equal):
    try:
        if Reference.split("<>")[0] in self.correspondences:
            DocPath = self.correspondences[Reference.split("<>")[0]]
        else:
            DocPath = Reference.split("<>")[0]

        DocWorkbook = openpyxl.load_workbook(DocPath, data_only=True)
        SheetName = Reference.split("<>")[1]
        DocCel = Reference.split("<>")[2]
        DocSheet = DocWorkbook[SheetName]
        CelValue = DocSheet[DocCel].value

        input_values = []
        try:
            column = int(Column)
            WorkSheet = self.Workbook['Config']
            for row in range(2, WorkSheet.max_row + 1):
                if WorkSheet.cell(row, int(Column)).value is not None:
                    input_values.append(WorkSheet.cell(row, int(Column)).value)
        except:
            input_values.append(Column)

        col = ""
        row = ""
        for char in DocCel:
            if char.isalpha():
                col += char
            else:
                row += char

        row = int(row)
        number_points = 0

        if str(Equal).casefold() == "true":
            while (DocSheet[col + str(row)].value is not None):
                if DocSheet[col + str(row)].value in input_values:
                    number_points += 1
                row += 1

        elif str(Equal).casefold() == "false":
            while (DocSheet[col + str(row)].value is not None):
                if DocSheet[col + str(row)].value not in input_values:
                    number_points += 1
                row += 1

        return number_points
    except:
        return 'NA'

def CheckMultipleValues(self, Column, Reference, Equal):

    try:
        WorkSheet = self.Workbook['Config']
        column = int(Column)
        list_values = []
        for row in range(2, WorkSheet.max_row + 1):
            if WorkSheet.cell(row, column).value is not None:
                list_values.append(WorkSheet.cell(row, column).value)

        if Reference.split("<>")[0] in self.correspondences:
            DocPath = self.correspondences[Reference.split("<>")[0]]
        else:
            DocPath = Reference.split("<>")[0]
        DocWorkbook = openpyxl.load_workbook(DocPath, data_only=True)
        current_results = []
        for sheet in self.vsm_sheets:
            DocSheet = DocWorkbook[sheet]
            Value = DocSheet[Reference.split("<>")[1]].value

            if str(Equal).casefold() == "true":
                if Value in list_values:
                    current_results.append('OK')
                else:
                    current_results.append('NOK')
            elif str(Equal).casefold() == "false":
                if Value not in list_values:
                    current_results.append('OK')
                else:
                    current_results.append('NOK')
        self.multiple_results.append(current_results)
    except:
        return 'NA'