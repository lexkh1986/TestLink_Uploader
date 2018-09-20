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

    def loadTestCases(self):
        iExcel = self.TEMPLATE
        for ir in iExcel.INDEX_LIST:
            newTC = TestCase()
            newTC.WbIndex = ir
            newTC.Sync = self.INFO.SYNC.get(iExcel.cell_byCol(ir, 'Sync').lower())
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
        isReadonly(self.FILEPATH)
        newPulled = self.INFO.pullTestCases()
        self.TEMPLATE.prepare_write()
        for iTC in self.INFO.TESTS:
            #Write pulled item to row
            for val in self.TEMPLATE.HEADER:
                iStyle = self.TEMPLATE.iCommonStyle
                iValue = parse_steps(iTC.Steps, True) if val == 'Steps' else iTC.toDict().get(val)
                if val == 'Sync':
                    iStyle = self.TEMPLATE.iCentreStyle
                    iValue = dict_getkey(self.INFO.SYNC, iValue)
                if iTC.fmtCode == 'iConflict':
                    iStyle = iStyle + ';pattern: pattern solid, fore_colour yellow'
                self.TEMPLATE.write(iTC.WbIndex, val, parse_summary(iValue, True), iStyle)

            if iTC.FullID in newPulled:
                print 'Pulled new: %s - %s' % (iTC.FullID, parse_summary(iTC.Name, True))
                continue
            elif iTC.Sync in (1, 2) and iTC.fmtCode == 'iConflict':
                print 'Conflict: %s - %s' % (iTC.FullID, parse_summary(iTC.Name, True))
                continue
            elif iTC.Sync not in (1, 2):
                print 'Overwrited: %s - %s' % (iTC.FullID, parse_summary(iTC.Name, True))
                continue
            
        self.TEMPLATE.save_write(self.FILEPATH)

    def pushTestCases(self):
        isReadonly(self.FILEPATH)
        self.TEMPLATE.prepare_write()
        for iTC in self.INFO.TESTS:
            if iTC.Sync == 1:
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
