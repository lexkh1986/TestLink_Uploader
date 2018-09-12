from Connection import *
from TestModel import *
from Misc import *
from Template import Template
from xlutils.copy import copy
import xlwt, xlrd
import os, sys

class Workbook(object):
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'
    CONFIG_PATH = '%s/../setting.txt' % os.path.dirname(__file__)
    
    def __init__(self):
        self.isLoaded = False
        self._setConfig()

    def _setConfig(self):
        if os.path.isfile(self.CONFIG_PATH):
            iConfig = getVarFromFile(self.CONFIG_PATH)
            self.USER = iConfig.USERNAME
            self.DEVKEY = iConfig.DEVKEY
            self.FILEPATH = iConfig.EXCEL_PATH
            self.SHEETNAME = iConfig.SHEET_NAME
            self.ISSTEPS = iConfig.SUMMARY_AS_STEP

    def _loadSheet(self):
        if os.path.isfile(self.FILEPATH):
            self.WORKBOOK = xlrd.open_workbook(self.FILEPATH, on_demand = False, formatting_info = True)
            self.SHEET = self.WORKBOOK.sheet_by_name(self.SHEETNAME)
            self.isLoaded = True
            return
        raise Exception('Could not locate excel workbook by path: %s' % self.FILEPATH)

    def _loadInfo(self):
        if not self.isLoaded:
            raise Exception('Workbook content is empty. Please make sure excel file is accessible')
        iTemplate = Template()
        self.INFO = Connection(self.DEVKEY,
                               self.SHEET.cell_value(iTemplate.LOC_PROJECT_VAL.get('r'), iTemplate.LOC_PROJECT_VAL.get('c')),
                               self.SHEET.cell_value(iTemplate.LOC_PLAN_VAL.get('r'), iTemplate.LOC_PLAN_VAL.get('c')),
                               self.SHEET.cell_value(iTemplate.LOC_BUILD_VAL.get('r'), iTemplate.LOC_BUILD_VAL.get('c')))

    def _pullTestCases(self):
        wb = copy(self.WORKBOOK)
        ws = wb.get_sheet(0)
        ws.show_grid = False

        iTemplate = Template()
        self.INFO.pullTestCases()
        iRow = iTemplate.LOC_DETAILS.get('r')
        iHeader = self.SHEET.row_values(iRow)
        
        for id, iTC in self.INFO.TESTS.items():
            iRow = iRow + 1
            for i, val in enumerate(iHeader):
                iStyle = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin;')
                if self.ISSTEPS and val == 'Steps':
                    ws.write(iRow, i, parse_summary('%s\n%s' % (iTC.toDict().get('Summary'), iTC.toDict().get('Steps')), True), iStyle)
                else:
                    ws.write(iRow, i, iTC.toDict().get(val), iStyle)
                    
        wb.save(self.FILEPATH)

    def _syncDetails(self):
        try:
            wb = copy(self.WORKBOOK)
            ws = wb.get_sheet(0)
            ws.show_grid = False
            wb.save(self.FILEPATH)
        except Exception, err:
            raise Exception(err)

if __name__ == "__main__":
   wb = Workbook()
   wb._loadSheet()
   print 'Workbook loaded (%s)\nSheet: %s\n' % (wb.FILEPATH, wb.SHEETNAME)
   
   wb._loadInfo()
   print 'Test Project: %s' % wb.INFO.PROJECT_NAME
   print 'Test Plan: %s' % wb.INFO.TESTPLAN_NAME
   print 'Test Build: %s' % wb.INFO.TESTBUILD_NAME

   if '-p' in sys.argv:
       wb._pullTestCases()
       print 'Data pulled from TestLink.... Done!'

   if '-s' in sys.argv:
       wb._syncDetails()
       print 'Data sync to TestLink.... Done!'
