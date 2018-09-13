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

    def _loadWorkbook(self):
        iTemplate = Template()
        if os.path.isfile(self.CONFIG_PATH):
            iConfig = getVarFromFile(self.CONFIG_PATH)
            self.USER = iConfig.USERNAME
            self.FILEPATH = iConfig.EXCEL_PATH
            if os.path.isfile(self.FILEPATH):
                self.WORKBOOK = xlrd.open_workbook(self.FILEPATH, on_demand = False, formatting_info = True)
                self.SHEET = self.WORKBOOK.sheet_by_name(iConfig.SHEET_NAME)
                self.HEADER = self.SHEET.row_values(iTemplate.LOC_DETAILS.get('r'))
                self.isLoaded = True

                self.INFO = Connection(iConfig.DEVKEY,
                                       self.SHEET.cell_value(iTemplate.LOC_PROJECT_VAL.get('r'), iTemplate.LOC_PROJECT_VAL.get('c')),
                                       self.SHEET.cell_value(iTemplate.LOC_PLAN_VAL.get('r'), iTemplate.LOC_PLAN_VAL.get('c')),
                                       self.SHEET.cell_value(iTemplate.LOC_BUILD_VAL.get('r'), iTemplate.LOC_BUILD_VAL.get('c')))
                self.INFO.DELIMETER = iConfig.TESTCASE_PATH_DELIMETER
                print 'Workbook loaded (%s)\nSheet: %s\n' % (self.FILEPATH, iConfig.SHEET_NAME)
                return
            raise Exception('Could not locate excel workbook by path: %s' % self.FILEPATH)
        raise Exception('Could not locate settings: %s' % self.CONFIG_PATH)

    def _loadTestCases(self):
        if not self.isLoaded:
            raise Exception('Workbook content is empty. Please make sure excel file is accessible')
        iTemplate = Template()
        iRow = iTemplate.LOC_DETAILS.get('r')

        for ir, row in enumerate(self.SHEET.get_rows()):
            iSyncFlag = self.SHEET.cell_value(ir, self.HEADER.index('Sync'))
            if ir <= iRow or not self.INFO.SYNC.get(iSyncFlag.upper(), False): continue
            newTC = TestCase()
            newTC.Sync = iSyncFlag
            newTC.WbIndex = ir
            newTC.FullID = self.SHEET.cell_value(ir, self.HEADER.index('FullID'))
            newTC.Name = self.SHEET.cell_value(ir, self.HEADER.index('Name'))
            newTC.Summary = self.SHEET.cell_value(ir, self.HEADER.index('Summary'))
            newTC.Address = self.SHEET.cell_value(ir, self.HEADER.index('Address'))
            newTC.Steps = self.SHEET.cell_value(ir, self.HEADER.index('Steps'))
            newTC.Author = self.SHEET.cell_value(ir, self.HEADER.index('Author')).lower()
            newTC.Priority = self.SHEET.cell_value(ir, self.HEADER.index('Priority')).title()
            newTC.Exectype = self.SHEET.cell_value(ir, self.HEADER.index('Exectype')).title()
            self.INFO.append_Test(newTC)

    def _pullTestCases(self):
        isFound = self.INFO.pullTestCases()
        if isFound:
            wb = copy(self.WORKBOOK)
            ws = wb.get_sheet(0)
            ws.show_grid = False
            iStyle = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin;')
            
            for iTC in self.INFO.TESTS:
                if iTC.SyncStatus != 1: continue
                for i, val in enumerate(self.HEADER):
                    ws.write(iTC.WbIndex, i, parse_summary(iTC.toDict().get(val), True), iStyle)
                print 'Pulled TestCase: %s - %s' % (iTC.FullID, iTC.Name)
            try:
                wb.save(self.FILEPATH)
            except IOError, err:
                raise Exception('Permission denied: %s\nPlease close your workbook and re-run task again.' % self.FILEPATH)
        else:
            print 'Testplan details pulled without any TestCases from TestLink'

    def _pushTestCases(self):
        try:
            wb = copy(self.WORKBOOK)
            ws = wb.get_sheet(0)
            ws.show_grid = False

            iStyle = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin;')
            for iTC in self.INFO.TESTS:
                if self.INFO.SYNC.get(iTC.Sync, False):
                    self.INFO.pushTestCase(iTC)
                    if iTC.SyncStatus == 3:
                        ws.write(iTC.WbIndex, self.HEADER.index('FullID'), iTC.FullID, iStyle)
                        print 'Successfully created TestCase: %s - %s' % (iTC.FullID, iTC.Name)
                    elif iTC.SyncStatus == 2:
                        print 'Successfully modified TestCase: %s - %s' % (iTC.FullID, iTC.Name)

            wb.save(self.FILEPATH)
        except IOError, err: raise Exception('Permission denied: %s\nPlease close your workbook and re-run task again.' % self.FILEPATH)
