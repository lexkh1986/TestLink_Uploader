from API_Excel import *
from Misc import *
from TestModel import *
from testlink import *
import sys, traceback

class Connection(Test):
    SYNC = {'':0, 'p':1, 'e':2}
    STATUS = {'p':'Pass', 'f':'Fail', 'n':'Not run', 'b':'Block'}
    STATE = {2:'Ready', 4:'Rework', 3:'Final', 1:'Draft'}
    IMPORTANCE = {3:'High', 2:'Medium', 1:'Low'}
    EXECUTIONTYPE = {1:'Manual', 2:'Automated'}

    def __init__(self, devkey, project, plan, build):
        super(Connection, self).__init__(project, plan, build)
        self.SERVER_URL = 'http://testlink.nexcel.vn/lib/api/xmlrpc/v1/xmlrpc.php'
        self.DEVKEY = devkey
        self.CONN = TestLinkHelper(self.SERVER_URL, self.DEVKEY).connect(TestlinkAPIGeneric)
        self._project()

    def _project(self):
        for elem in self.CONN.getProjects():
            if rem_empty(elem['name']) == rem_empty(self.PROJECT_NAME):
                self.PROJECT_ID = elem['id']
                self.PROJECT_PREFIX = elem['prefix']
                return
        raise Exception('Project name not found: %s' % self.PROJECT_NAME)

    def _testplan(self):
        try:
            tmpName = rem_empty(self.TESTPLAN_NAME)
            tmpFound = [(i['name'], i['id']) for i in self.CONN.getProjectTestPlans(self.PROJECT_ID)]
            for tmpTP in tmpFound:
                if tmpName == tmpTP[0].replace(' ',''):
                    self.TESTPLAN_ID = tmpTP[1]
                    return True
        except Exception, err:
            print 'TestPlan name not found: %s' % self.TESTPLAN_NAME
            sys.exit(1)

    def _testbuild(self):
        iBuilds = self.CONN.getBuildsForTestPlan(self.TESTPLAN_ID)
        for i in iBuilds:
            if rem_empty(self.TESTBUILD_NAME) == rem_empty(i['name']):
                self.TESTBUILD_ID = i['id']
                return True
        print 'Test build not found: %s' % self.TESTBUILD_NAME
        sys.exit(1)

    def _getTestCase_byID(self, full_external_id):
        return self.CONN.getTestCase(testcaseexternalid = full_external_id)

    def _getTestCase_byName(self, name):
        try:
            tmpTC = self.CONN.getTestCaseIDByName(testcasename = name,
                                                  testprojectname = self.PROJECT_NAME)
            return tmpTC
        except: return []

    def _addTestCase_toTestPlan(self, iTC_):
        self.CONN.addTestCaseToTestPlan(testprojectid = self.PROJECT_ID,
                                        testplanid = self.TESTPLAN_ID,
                                        testcaseexternalid = iTC_.FullID,
                                        version = iTC_.Version,
                                        overwrite = True)

    def _addTestCase_toTestBuild(self, iTC_):
        self.CONN.assignTestCaseExecutionTask(user = iTC_.Owner,
                                              testplanid = self.TESTPLAN_ID,
                                              buildid = self.TESTBUILD_ID,
                                              testcaseexternalid = iTC_.FullID)

    def _removeTestCase_fromTestBuild(self, iTC_):
        self.CONN.unassignTestCaseExecutionTask(action = 'unassignAll',
                                                testplanid = self.TESTPLAN_ID,
                                                buildid = self.TESTBUILD_ID,
                                                testcaseexternalid = iTC_.FullID)

    def _getTestCase_Owner(self, iTC_):
        return self.CONN.getTestCaseAssignedTester(testplanid = self.TESTPLAN_ID,
                                                   buildid = self.TESTBUILD_ID,
                                                   testcaseexternalid = iTC_.FullID)

    def _createTestPlan(self):
        try:
            rs = self.CONN.createTestPlan(testprojectname = self.PROJECT_NAME,
                                          testplanname = self.TESTPLAN_NAME)
            self.TESTPLAN_ID = rs[0]['id']
            print 'Created TestPlan: %s' % self.TESTPLAN_NAME
        except Exception, err:
            self._testplan()
            print 'Failed to create TestPlan: %s. Please recheck if already exists' % self.TESTPLAN_NAME

    def _createTestBuild(self):
        rs = self.CONN.createBuild(testplanid = self.TESTPLAN_ID,
                                   buildname = self.TESTBUILD_NAME)
        self.TESTBUILD_ID = rs[0]['id']
        if rs[0]['status']:
            print 'Created TestBuild: %s' % self.TESTBUILD_NAME
        else:
            print 'Failed to create TestBuild: %s. Please recheck if already exists' % self.TESTBUILD_NAME

    def _getFullSuitePath(self, full_external_id):
        tmpPath = None
        if full_external_id is not None:
            ID = self._getTestCase_byID(full_external_id)[0]['testcase_id']
            try:
                tmpPath = self.CONN.getFullPath(int(ID)).values()
                return self.DELIMETER.join(tmpPath[0][1:])
            except Exception, err:
                traceback.print_exc()
        return tmpPath

    def _validateParentSuite(self, iTC_):
        tmpPath = iTC_.Address.split(self.DELIMETER)
        tmpRefID, tmpRefName = None, None
        tmpRoot = self.CONN.getFirstLevelTestSuitesForTestProject(self.PROJECT_ID)

        # Get first child
        for fc in tmpRoot:
            if fc['name'] == tmpPath[0]:
                tmpRefID, tmpRefName = fc['id'], fc['name']
        if tmpRefID is None: raise Exception('Could not found "%s" in full path: %s' % (tmpPath[0], iTC_.Address))

        # Get child loop
        for idx, node in enumerate(tmpPath):
            if idx == 0: continue
            tmpChilds = self.CONN.getTestSuitesForTestSuite(tmpRefID)
            if not tmpChilds: raise Exception('Could not found "%s" in full path: %s' % (node, iTC_.Address))
            tmpChilds = tmpChilds.values() if tmpChilds.get('name', None) is None else [tmpChilds]
            for n in tmpChilds:
                if n['name'] == node:
                    tmpRefID, tmpRefName = n['id'], n['name']
                    break
        return tmpRefID

    def pushResult(self, iTC_):
        try:
            iResult = dict_getkey(self.STATUS, self.DEFAULT_RESULT)
            if not self.USE_DEFAULT_RESULT:
                iResult = dict_getkey(self.STATUS, iTC_.Result)
                if iResult in ('n', None):
                    return [{'status':False, 'message': 'skipped'}]
            return self.CONN.reportTCResult(testplanid = self.TESTPLAN_ID,
                                            testcaseexternalid = iTC_.FullID,
                                            buildid = self.TESTBUILD_ID,
                                            status = iResult,
                                            user = iTC_.Owner,
                                            notes = iTC_.Note,
                                            overwrite = True)
        except Exception, err:
            return [{'status':False, 'message':err}]

    def pushTestCase(self, iTC_):
        #Check if already exists before
        iDupList = self._getTestCase_byName(iTC_.Name)
        if iDupList:
            for elem in iDupList:
                if elem['parent_id'] == self._validateParentSuite(iTC_) and iTC_.ID <> elem['tc_external_id']:
                    print 'A duplicate name found at row %s (with %s-%s: %s) in same folder. Please use another name'\
                          % (str(iTC_.WbIndex+1), self.PROJECT_PREFIX,
                             elem['tc_external_id'], elem['name'])
                    return 0

        #Check if missing owner (in case auto assigned to testplan)
        if self.AUTO_ADD_TESTPLAN:
            if iTC_.Owner in ('', None):
                print 'Missing execution owner at row %s (%s). Please update your workbook' % (str(iTC_.WbIndex+1), iTC_.Name)
                return 0

        #Create if new TestCase
        if iTC_.FullID in ('', None):
            try:
                rs = self.CONN.createTestCase(testcasename = iTC_.Name,
                                              testprojectid = self.PROJECT_ID,
                                              testsuiteid = self._validateParentSuite(iTC_),
                                              authorlogin = iTC_.Author,
                                              summary = parse_summary(iTC_.Summary),
                                              steps = iTC_.Steps,
                                              importance = dict_getkey(self.IMPORTANCE, iTC_.Priority),
                                              executiontype = dict_getkey(self.EXECUTIONTYPE, iTC_.Exectype))
                iTC_.FullID = '%s-%s' % (self.PROJECT_PREFIX, rs[0]['additionalInfo']['external_id'])
                if self.AUTO_ADD_TESTPLAN:
                    self._addTestCase_toTestPlan(iTC_)
                    self._addTestCase_toTestBuild(iTC_)
                print 'Successfully created TestCase: (Row %s) %s - %s' % (str(iTC_.WbIndex+1), iTC_.FullID, iTC_.Name)
                return 1
            except Exception, err:
                print 'Failed to create TestCase: (Row %s) %s\n%s' % (str(iTC_.WbIndex+1), iTC_.Name, err)

        #Modify existing TestCase
        if iTC_.FullID not in ('', None):
            try:
                if self.IGNORE_PUSH_STEPS:
                    self.CONN.updateTestCase(testcaseexternalid = iTC_.FullID,
                                             testcasename = iTC_.Name,
                                             user = iTC_.Author,
                                             importance = dict_getkey(self.IMPORTANCE, iTC_.Priority),
                                             executiontype = dict_getkey(self.EXECUTIONTYPE, iTC_.Exectype))
                else:
                    self.CONN.updateTestCase(testcaseexternalid = iTC_.FullID,
                                             testcasename = iTC_.Name,
                                             user = iTC_.Author,
                                             summary = parse_summary(iTC_.Summary),
                                             steps = iTC_.Steps,
                                             importance = dict_getkey(self.IMPORTANCE, iTC_.Priority),
                                             executiontype = dict_getkey(self.EXECUTIONTYPE, iTC_.Exectype))
                if self.AUTO_ADD_TESTPLAN:
                    self._addTestCase_toTestPlan(iTC_)
            except Exception, err:
                if type(err).__name__ == 'ExpatError': pass
                else:
                    print 'Failed to modifiy TestCase: (Row %s) %s\n%s' % (str(iTC_.WbIndex+1), iTC_.Name, err)
                    return 0
                
            #Reassign owner
            if self.AUTO_ADD_TESTPLAN:
                try:
                    self._removeTestCase_fromTestBuild(iTC_)
                    self._addTestCase_toTestBuild(iTC_)
                except Exception, err:
                    if type(err).__name__ == 'ExpatError': pass
                    else:
                        print 'Failed to reassign owner for TestCase: (Row %s) %s\n%s' % (str(iTC_.WbIndex+1), iTC_.Name, err)
                        return 0
                    
            print 'Successfully modified TestCase: (Row %s) %s - %s' % (str(iTC_.WbIndex+1), iTC_.FullID, iTC_.Name)
            return 2

    def pullTestCases(self):
        iTemplate = Template()
        iNewTCs = []
        iStart = iTemplate.LOC_DETAILS.get('r') + 1
        iConsumedLoc = [elem.toDict()['WbIndex'] for elem in self.TESTS]
        iOpenLoc = [i for i in range(iStart, iTemplate.MAX_ROW_SUPPORT) if i not in iConsumedLoc]
        iExistingFullID = [iTC.FullID for iTC in self.TESTS]
        iTC_List = self.CONN.getTestCasesForTestPlan(testplanid = self.TESTPLAN_ID,
                                                     buildid = self.TESTBUILD_ID,
                                                     details = 'simple')
        if iTC_List:
            iTC_List = iTC_List.values()
            for iTC in iTC_List:
                iTC_Details = self._getTestCase_byID(iTC[0]['full_external_id'])
                newTC = TestCase()
                newTC.Sync = 0
                newTC.ID = iTC[0]['external_id']
                newTC.FullID = iTC[0]['full_external_id']
                newTC.Name = iTC_Details[0]['name']
                newTC.Address = self._getFullSuitePath(newTC.FullID)
                newTC.Summary = iTC_Details[0]['summary']
                newTC.Steps = iTC_Details[0]['steps']
                newTC.Author = iTC_Details[0]['author_login'].lower()
                newTC.Owner = self._getTestCase_Owner(newTC)[0]['login'].lower()
                newTC.Version = int(iTC_Details[0]['version'])
                newTC.Priority = self.IMPORTANCE.get(int(iTC_Details[0]['importance']))
                newTC.Exectype = self.EXECUTIONTYPE.get(int(iTC_Details[0]['execution_type']))
                if iTC[0]['exec_status'] != 'n':
                    newTC.Result = self.STATUS.get(iTC[0]['exec_status'])
                    newTC.Duration = iTC[0]['execution_duration']

                #Handle conflict
                if newTC.FullID not in iExistingFullID:
                    iTakenLoc = iOpenLoc.pop(0)
                    newTC.WbIndex = iTakenLoc
                    self.append_Test(newTC)
                    iNewTCs.append(newTC.FullID)
                else:
                    currTCs = self.get_byFullID(iTC[0]['full_external_id'])
                    if currTCs.Sync not in (1, 2):
                        newTC.WbIndex = currTCs.WbIndex
                        self.pop_byFullID(currTCs.FullID)
                        self.append_Test(newTC)
                    else:
                        tmpCmp = {'FullID':currTCs.FullID == newTC.FullID,
                                  'Address':currTCs.Address == newTC.Address,
                                  'Name':currTCs.Name == newTC.Name,
                                  'Summary':rem_empty(currTCs.Summary) == rem_empty(parse_summary(newTC.Summary, True)),
                                  'Author':currTCs.Author == newTC.Author,
                                  'Owner':currTCs.Owner == newTC.Owner,
                                  'Priority':currTCs.Priority == newTC.Priority,
                                  'Exectype':currTCs.Exectype == newTC.Exectype}
                        if False in tmpCmp.values():
                            currTCs.fmtCode = tmpCmp
        return iNewTCs

    def connectTL(self):
        self._testplan()
        self._testbuild()
