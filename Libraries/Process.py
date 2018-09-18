from API_Excel import *
from API_TestLink import *
from TestModel import *
from Misc import *
import os, sys

class Workbook(object):
    ROBOT_LIBRARY_SCOPE = 'GLOBAL'
    CONFIG_PATH = '%s/../setting.txt' % os.path.dirname(__file__)
    
    def __init__(self):
        self.TEMPLATE = Template()

    def loadWorkbook(self):
        iExcel = self.TEMPLATE
        if os.path.isfile(self.CONFIG_PATH):
            iConfig = getVarFromFile(self.CONFIG_PATH)
            self.USER = iConfig.USERNAME
            self.FILEPATH = iConfig.EXCEL_PATH
            iExcel.load(iConfig.EXCEL_PATH, iConfig.SHEET_NAME)
            self.INFO = Connection(iConfig.DEVKEY, iExcel.get_ProjectName(), iExcel.get_TestPlanName(), iExcel.get_TestBuildName())
            self.INFO.DELIMETER = iConfig.TESTCASE_PATH_DELIMETER
            self.INFO.AUTO_ADD_TESTPLAN = iConfig.AUTO_ADD_TESTPLAN
            self.INFO.USE_DEFAULT_RESULT = iConfig.USE_DEFAULT_RESULT
            self.INFO.DEFAULT_RESULT = iConfig.DEFAULT_RESULT_IN_BATCH
            return
        raise Exception('Could not locate settings: %s' % self.CONFIG_PATH)

    def loadTestCases(self, useSyncFlag=True):
        iExcel = self.TEMPLATE
        for ir in iExcel.INDEX_LIST:
            iSyncFlag = iExcel.cell_byCol(ir, 'Sync')
            if useSyncFlag:
                if not self.INFO.SYNC.get(iSyncFlag.upper(), False): continue
            newTC = TestCase()
            newTC.Sync = iSyncFlag
            newTC.WbIndex = ir
            newTC.FullID = iExcel.cell_byCol(ir, 'FullID')
            newTC.Name = iExcel.cell_byCol(ir, 'Name')
            newTC.Summary = iExcel.cell_byCol(ir, 'Summary')
            newTC.Address = iExcel.cell_byCol(ir, 'Address')
            newTC.Author = iExcel.cell_byCol(ir, 'Author').lower()
            newTC.Owner = iExcel.cell_byCol(ir, 'Owner').lower()
            newTC.Priority = iExcel.cell_byCol(ir, 'Priority').title()
            newTC.Exectype = iExcel.cell_byCol(ir, 'Exectype').title()
            newTC.Result = iExcel.cell_byCol(ir, 'Result')
            newTC.Duration = iExcel.cell_byCol(ir, 'Duration')
            newTC.Note = iExcel.cell_byCol(ir, 'Note')
            newTC.ID = newTC.FullID[len(self.INFO.PROJECT_PREFIX)+1-len(newTC.FullID):]
            if 'Steps' in iExcel.HEADER:
                newTC.Steps = parse_steps(iExcel.cell_byCol(ir, 'Steps'))
            self.INFO.append_Test(newTC)

    def pullTestCases(self):
        pulledList = self.INFO.pullTestCases()
        if pulledList:
            self.TEMPLATE.prepare_write()
            for iTC in self.INFO.TESTS:
                if iTC.FullID in pulledList:
                    for val in self.TEMPLATE.HEADER:
                        iStyle = self.TEMPLATE.iCentreStyle if val in ('Sync', 'Result') else self.TEMPLATE.iCommonStyle
                        iValue = parse_steps(iTC.Steps, True) if val == 'Steps' else iTC.toDict().get(val)
                        self.TEMPLATE.write(iTC.WbIndex, val, parse_summary(iValue, True), iStyle)
                    print 'Pulled TestCase: %s - %s' % (iTC.FullID, parse_summary(iTC.Name, True))
            self.TEMPLATE.save_write(self.FILEPATH)
        else:
            print 'Testplan details pulled without any TestCases from TestLink'

    def pushTestCases(self):
        self.TEMPLATE.prepare_write()
        for iTC in self.INFO.TESTS:
            if self.INFO.SYNC.get(iTC.Sync, False):
                if self.INFO.pushTestCase(iTC) == 1:
                    self.TEMPLATE.write(iTC.WbIndex, 'FullID', iTC.FullID, self.TEMPLATE.iCommonStyle)
        self.TEMPLATE.save_write(self.FILEPATH)

    def pushResults(self):
        for iTC in self.INFO.TESTS:
            if iTC.FullID not in ('', None):
                rs = self.INFO.pushResult(iTC)
                if rs[0]['status']:
                    print 'Succeeded upload result: %s %s' % (iTC.FullID, iTC.Result)
                else:
                    print 'Execution result: %s %s - %s' % (iTC.FullID, iTC.Result, rs[0]['message'])

    def connect(self):
        self.INFO.connectTL()

    def createTestplan(self):
        self.INFO._createTestPlan()
        self.INFO._createTestBuild()
