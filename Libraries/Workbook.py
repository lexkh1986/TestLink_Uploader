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

    def loadWorkbook(self):
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
                self.INFO.AUTO_ADD_TESTPLAN = iConfig.AUTO_ADD_TESTPLAN
                self.INFO.USE_DEFAULT_RESULT = iConfig.USE_DEFAULT_RESULT
                self.INFO.DEFAULT_RESULT = iConfig.DEFAULT_RESULT_IN_BATCH
                print 'Workbook loaded (%s)\nSheet: %s\n' % (self.FILEPATH, iConfig.SHEET_NAME)
                return
            raise Exception('Could not locate excel workbook by path: %s' % self.FILEPATH)
        raise Exception('Could not locate settings: %s' % self.CONFIG_PATH)

    def loadTestCases(self, useSyncFlag=True):
        if not self.isLoaded:
            raise Exception('Workbook content is empty. Please make sure excel file is accessible')
        iTemplate = Template()
        iRow = iTemplate.LOC_DETAILS.get('r')

        for ir, row in enumerate(self.SHEET.get_rows()):
            if ir > iRow:
                if False not in list(set([self.SHEET.cell_type(ir, i) in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK)\
                                          for i, cell in enumerate(row)])): continue
                iSyncFlag = self.SHEET.cell_value(ir, self.HEADER.index('Sync'))
                if useSyncFlag:
                    if not self.INFO.SYNC.get(iSyncFlag.upper(), False): continue
                newTC = TestCase()
                newTC.Sync = iSyncFlag
                newTC.WbIndex = ir
                newTC.FullID = self.SHEET.cell_value(ir, self.HEADER.index('FullID'))
                newTC.Name = self.SHEET.cell_value(ir, self.HEADER.index('Name'))
                newTC.Summary = self.SHEET.cell_value(ir, self.HEADER.index('Summary'))
                newTC.Address = self.SHEET.cell_value(ir, self.HEADER.index('Address'))
                newTC.Steps = self.SHEET.cell_value(ir, self.HEADER.index('Steps'))
                newTC.Author = self.SHEET.cell_value(ir, self.HEADER.index('Author')).lower()
                newTC.Owner = self.SHEET.cell_value(ir, self.HEADER.index('Owner'))
                newTC.Priority = self.SHEET.cell_value(ir, self.HEADER.index('Priority')).title()
                newTC.Exectype = self.SHEET.cell_value(ir, self.HEADER.index('Exectype')).title()

                newTC.Result = self.SHEET.cell_value(ir, self.HEADER.index('Result'))
                newTC.Duration = self.SHEET.cell_value(ir, self.HEADER.index('Duration'))
                newTC.Note = self.SHEET.cell_value(ir, self.HEADER.index('Note'))

                newTC.ID = newTC.FullID[len(self.INFO.PROJECT_PREFIX)+1-len(newTC.FullID):]
                self.INFO.append_Test(newTC)

    def pullTestCases(self):
        pulledList = self.INFO.pullTestCases()
        if pulledList:
            wb = copy(self.WORKBOOK)
            ws = wb.get_sheet(0)
            ws.show_grid = False
            iStyle = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin;')
            
            for iTC in self.INFO.TESTS:
                if iTC.FullID in pulledList:
                    for i, val in enumerate(self.HEADER):
                        ws.write(iTC.WbIndex, i, parse_summary(iTC.toDict().get(val), True), iStyle)
                    print 'Pulled TestCase: %s - %s' % (iTC.FullID, iTC.Name)
            try:
                wb.save(self.FILEPATH)
            except IOError, err:
                raise Exception('Permission denied: %s\nPlease close your workbook and re-run task again.' % self.FILEPATH)
        else:
            print 'Testplan details pulled without any TestCases from TestLink'

    def pushTestCases(self):
        try:
            wb = copy(self.WORKBOOK)
            ws = wb.get_sheet(0)
            ws.show_grid = False

            iStyle = xlwt.easyxf('borders: left thin, right thin, top thin, bottom thin;')
            for iTC in self.INFO.TESTS:
                if self.INFO.SYNC.get(iTC.Sync, False):
                    if self.INFO.pushTestCase(iTC) == 1:
                        ws.write(iTC.WbIndex, self.HEADER.index('FullID'), iTC.FullID, iStyle)

            wb.save(self.FILEPATH)
        except IOError, err: raise Exception('Permission denied: %s\nPlease close your workbook and re-run task again.' % self.FILEPATH)

    def pushResults(self):
        for iTC in self.INFO.TESTS:
            if iTC.FullID not in ('', None):
                rs = self.INFO.pushResult(iTC)
                if rs[0]['status']:
                    print 'Succeeded upload result: %s %s' % (iTC.FullID, iTC.Result)
                else:
                    print 'Execution result: %s %s - %s' % (iTC.FullID, iTC.Result, rs[0]['message'])
