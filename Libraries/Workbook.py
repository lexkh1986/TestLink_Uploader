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
        self.INFO = Test(self.SHEET.cell_value(iTemplate.LOC_PROJECT_VAL.get('r'), iTemplate.LOC_PROJECT_VAL.get('c')),
                         self.SHEET.cell_value(iTemplate.LOC_PLAN_VAL.get('r'), iTemplate.LOC_PLAN_VAL.get('c')),
                         self.SHEET.cell_value(iTemplate.LOC_BUILD_VAL.get('r'), iTemplate.LOC_BUILD_VAL.get('c')))

    def _saveDetails(self):
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
   
   if '-s' in sys.argv:
       wb._saveDetails()
       print 'Data sync to TestLink.... Done!'
