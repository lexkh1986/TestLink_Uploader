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
            iTemplate = Template()
            self.WORKBOOK = xlrd.open_workbook(self.FILEPATH, on_demand = False, formatting_info = True)
            self.SHEET = self.WORKBOOK.sheet_by_name(self.SHEETNAME)
            self.HEADER = self.SHEET.row_values(iTemplate.LOC_DETAILS.get('r'))
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

    def _loadTestCases(self):
        if not self.isLoaded:
            raise Exception('Workbook content is empty. Please make sure excel file is accessible')
        iTemplate = Template()
        iRow = iTemplate.LOC_DETAILS.get('r')

        for ir, row in enumerate(self.SHEET.get_rows()):
            iSyncFlag = bool(self.SHEET.cell_value(ir, self.HEADER.index('Sync')))
            if ir <= iRow or not iSyncFlag: continue
            newTC = TestCase()
            newTC.Sync = iSyncFlag
            newTC.WbIndex = ir
            newTC.FullID = self.SHEET.cell_value(ir, self.HEADER.index('FullID'))
            newTC.Name = self.SHEET.cell_value(ir, self.HEADER.index('Name'))
            newTC.Summary = self.SHEET.cell_value(ir, self.HEADER.index('Summary'))
            newTC.Address = self.SHEET.cell_value(ir, self.HEADER.index('Address'))
            newTC.Steps = self.SHEET.cell_value(ir, self.HEADER.index('Steps'))
            newTC.Author = self.SHEET.cell_value(ir, self.HEADER.index('Author'))
            newTC.Priority = self.SHEET.cell_value(ir, self.HEADER.index('Priority'))
            newTC.Exectype = self.SHEET.cell_value(ir, self.HEADER.index('Exectype'))
            self.INFO.append_Test(newTC)

    def _pullTestCases(self):
        self.INFO.pullTestCases()
        wb = copy(self.WORKBOOK)
        ws = wb.get_sheet(0)
        ws.show_grid = False

        iStyle = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin;')
        for ir, iTC in enumerate(self.INFO.TESTS):
            if iTC.FullID in ('', None): continue
            for i, val in enumerate(self.HEADER):
                if val in ('Summary', 'Step'):
                    ws.write(iTC.WbIndex, i, parse_summary(iTC.toDict().get(val), True), iStyle)
                else:
                    ws.write(iTC.WbIndex, i, iTC.toDict().get(val), iStyle)
        try:
            wb.save(self.FILEPATH)
        except IOError, err: raise Exception('Permission denied: %s\nPlease close your workbook and re-run task again.' % self.FILEPATH)

    def _syncDetails(self):
        try:
            self.INFO.pushTestCases()
            wb = copy(self.WORKBOOK)
            ws = wb.get_sheet(0)
            ws.show_grid = False

            iStyle = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin;')
            for iTC in self.INFO.TESTS:
                if iTC.Sync:
                    if self.SHEET.cell_value(iTC.WbIndex, self.HEADER.index('FullID')) in ('', None):
                        ws.write(iTC.WbIndex, self.HEADER.index('FullID'), iTC.FullID, iStyle)
                        print 'Successfully created TestCase: %s - %s' % (iTC.FullID, iTC.Name)
                    else:
                        print 'Successfully modified TestCase: %s - %s' % (iTC.FullID, iTC.Name)
            
            wb.save(self.FILEPATH)
        except IOError, err: raise Exception('Permission denied: %s\nPlease close your workbook and re-run task again.' % self.FILEPATH)

if __name__ == "__main__":
    wb = Workbook()
    wb._loadSheet()
    print 'Workbook loaded (%s)\nSheet: %s\n' % (wb.FILEPATH, wb.SHEETNAME)

    wb._loadInfo()
    print 'Test Project: %s' % wb.INFO.PROJECT_NAME
    print 'Test Plan: %s' % wb.INFO.TESTPLAN_NAME
    print 'Test Build: %s' % wb.INFO.TESTBUILD_NAME

    wb._loadTestCases()

    if set(sys.argv) & set(['-p','--pull']) or True:
        wb._pullTestCases()
        print 'Testplan details pulling from TestLink.... Done!'

    if set(sys.argv) & set(['-s','--sync']) or True:
        wb._syncDetails()
        print 'Testcases details syncing to TestLink.... Done!'
