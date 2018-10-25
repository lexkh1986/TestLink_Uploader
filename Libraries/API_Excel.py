from xlutils.copy import copy
from Misc import *
import xlwt, xlrd
import os, sys

class Template(object):
    LOC_PROJECT_LABLE = {'c':0, 'r':0}
    LOC_PROJECT_VAL = {'c':1, 'r':0}
    LOC_PLAN_LABLE = {'c':0, 'r':0}
    LOC_PLAN_VAL = {'c':1, 'r':1}
    LOC_BUILD_LABLE = {'c':0, 'r':2}
    LOC_BUILD_VAL = {'c':1, 'r':2}
    LOC_DETAILS = {'c':0, 'r':5}
    MAX_ROW_SUPPORT = 5000

    def __init__(self):
        self.isLoaded = False

    def initStyle(self):
        tmpFont = xlwt.Font()
        tmpFont.bold = False

        tmpBorder = xlwt.Borders()
        tmpBorder.left = 1
        tmpBorder.right = 1
        tmpBorder.top = 1
        tmpBorder.bottom = 1

        tmpAlignment = xlwt.Alignment()
        tmpAlignment.vert = xlwt.Alignment.VERT_CENTER
        tmpAlignment.wrap = xlwt.Alignment.WRAP_AT_RIGHT

        tmpPattern = xlwt.Pattern()

        iStyle = xlwt.XFStyle()
        iStyle.font = tmpFont
        iStyle.borders = tmpBorder
        iStyle.alignment = tmpAlignment
        iStyle.pattern = tmpPattern
        return iStyle

    def load(self, filepath, sheetname):
        self.openWorkbook(filepath)
        self.openSheet(sheetname)
        self.processData()
        print 'Workbook loaded (%s)\nSheet: %s\n' % (filepath, sheetname)

    def get_ProjectName(self):
        iVal = self.cell(self.LOC_PROJECT_VAL.get('r'),
                         self.LOC_PROJECT_VAL.get('c'))
        print 'Test Project: %s' % remove_endash(iVal)
        return remove_endash(iVal)

    def get_TestPlanName(self):
        iVal = self.cell(self.LOC_PLAN_VAL.get('r'),
                         self.LOC_PLAN_VAL.get('c'))
        print 'Test Plan: %s' % remove_endash(iVal)
        return remove_endash(iVal)

    def get_TestBuildName(self):
        iVal = self.cell(self.LOC_BUILD_VAL.get('r'),
                         self.LOC_BUILD_VAL.get('c'))
        print 'Test Build: %s' % remove_endash(iVal)
        return remove_endash(iVal)

    def openWorkbook(self, filepath):
        if not os.path.isfile(filepath):
            raise Exception('Could not locate excel workbook by path: %s' % filepath)
        self.WORKBOOK = xlrd.open_workbook(filepath, on_demand = False, formatting_info = True)

    def openSheet(self, name):
        try:
            self.SHEET = self.WORKBOOK.sheet_by_name(name)
            self.HEADER = self.SHEET.row_values(self.LOC_DETAILS.get('r'))
            self.isLoaded = True
        except Exception, err:
            raise Exception('Could not locate sheet name %s in workbook' % name)

    def processData(self):
        if not self.isLoaded:
            raise Exception('Workbook content is empty. Please make sure excel file is accessible')
        self.DATA = self.SHEET.get_rows()
        tmpIDXList = []
        for idx, row in enumerate(self.DATA):
            if idx > self.LOC_DETAILS.get('r')\
            and False in [self.SHEET.cell_type(idx, i) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK) for i, cell in enumerate(row)]:
                tmpIDXList.append(idx)
        self.INDEX_LIST = tmpIDXList

    def cell(self, row, col):
        return self.SHEET.cell_value(row, col)

    def cell_byCol(self, row, col_name):
        return self.SHEET.cell_value(row, self.HEADER.index(col_name))

    def prepare_write(self):
        self.tmpWORKBOOK = copy(self.WORKBOOK)
        self.tmpSHEET = self.tmpWORKBOOK.get_sheet(0)
        self.tmpSHEET.show_grid = False

    def save_write(self, filepath):
        try:
            self.tmpWORKBOOK.save(filepath)
        except IOError, err:
            print 'Permission denied: %s\nPlease close your workbook and re-run task again.' % filepath
            sys.exit(1)

    def write(self, row, col_name, value, style):
        self.tmpSHEET.write(row, self.HEADER.index(col_name), value, style)
