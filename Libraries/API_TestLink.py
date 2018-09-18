from API_Excel import *
from Misc import *
from TestModel import *
from testlink import *
import sys

class Connection(Test):
    SYNC = {'':False, 'X':True, 'x':True}
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
            if elem['name'] == self.PROJECT_NAME:
                self.PROJECT_ID = elem['id']
                self.PROJECT_PREFIX = elem['prefix']
                return
        raise Exception('Project name not found: %s' % self.PROJECT_NAME)

    def _testplan(self):
        try:
            self.TESTPLAN_ID = self.CONN.getTestPlanByName(testprojectname = self.PROJECT_NAME,
                                                       testplanname = self.TESTPLAN_NAME)[0]['id']
            return True
        except Exception, err:
            print 'TestPlan name not found: %s' % self.TESTPLAN_NAME
            sys.exit(1)

    def _testbuild(self):
        iBuilds = self.CONN.getBuildsForTestPlan(self.TESTPLAN_ID)
        for i in iBuilds:
            if self.TESTBUILD_NAME == i['name']:
                self.TESTBUILD_ID = i['id']
                return True
        print 'Test build not found: %s' % self.TESTBUILD_NAME
        sys.exit(1)

    def _getTestCase_byID(self, full_external_id):
        return self.CONN.getTestCase(testcaseexternalid = full_external_id)

    def _getTestCase_byName(self, name):
        try: return self.CONN.getTestCaseIDByName(testcasename = name,
                                                  testprojectname = self.PROJECT_NAME)
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

    def _getFullSuitePath(self, parentSuiteID):
        iParentDetails = self.CONN.getTestSuiteByID(parentSuiteID)
        iFullPath = [(iParentDetails['parent_id'], iParentDetails['name'])]

        while int(iFullPath[0][0]) != int(self.PROJECT_ID):
            tmpUpperDetails = self.CONN.getTestSuiteByID(iFullPath[0][0])
            iFullPath.insert(0, (tmpUpperDetails['parent_id'], tmpUpperDetails['name']))
        return self.DELIMETER.join([node[1] for node in iFullPath])

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
                    isFound = True
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
        if iDupList and iTC_.ID <> iDupList[0]['tc_external_id']:
            print 'A duplicate name found at row %s (with %s-%s: %s). Please use another name'\
                  % (iTC_.WbIndex, self.PROJECT_PREFIX,
                     iDupList[0]['tc_external_id'], iDupList[0]['name'])
            return 0

        #Check if missing owner (in case auto assigned to testplan)
        if self.AUTO_ADD_TESTPLAN:
            if iTC_.Owner in ('', None):
                print 'Missing execution owner at row %s. Please update your workbook' % iTC_.WbIndex
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
                print 'Successfully created TestCase: %s - %s' % (iTC_.FullID, iTC_.Name)
                return 1
            except Exception, err:
                print 'Failed to create TestCase: %s\n%s' % (iTC_.Name, err)

        #Modify existing TestCase
        if iTC_.FullID not in ('', None):
            try:
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
                    print 'Failed to modifiy TestCase: %s\n%s' % (iTC_.Name, err)
                    return 0
                
            #Reassign owner
            if self.AUTO_ADD_TESTPLAN:
                try:
                    self._removeTestCase_fromTestBuild(iTC_)
                    self._addTestCase_toTestBuild(iTC_)
                except Exception, err:
                    if type(err).__name__ == 'ExpatError': pass
                    else:
                        print 'Failed to reassign owner for TestCase: %s\n%s' % (iTC_.Name, err)
                        return 0
                    
            print 'Successfully modified TestCase: %s - %s' % (iTC_.FullID, iTC_.Name)
            return 2

    def pullTestCases(self):
        iTemplate = Template()
        iStart = iTemplate.LOC_DETAILS.get('r') + 1
        iConsumedLoc = [elem.toDict()['WbIndex'] for elem in self.TESTS]
        iOpenLoc = [i for i in range(iStart, iTemplate.MAX_ROW_SUPPORT) if i not in iConsumedLoc]
        iExistingFullID = [iTC.FullID for iTC in self.TESTS]
        iTC_List = self.CONN.getTestCasesForTestPlan(testplanid = self.TESTPLAN_ID,
                                                     buildid = self.TESTBUILD_ID,
                                                     details = 'simple')
        if iTC_List:
            iTC_List = iTC_List.values()
            tmpResults = []
            for iTC in iTC_List:
                iTC_Details = self._getTestCase_byID(iTC[0]['full_external_id'])
                newTC = TestCase()
                if iTC[0]['full_external_id'] in iExistingFullID:
                    newTC.WbIndex = self.get_byFullID(iTC[0]['full_external_id']).WbIndex
                    self.pop_byFullID(iTC[0]['full_external_id'])
                else:
                    iTakenLoc = iOpenLoc.pop(0)
                    newTC.WbIndex = iTakenLoc
                newTC.Sync = dict_getkey(self.SYNC, newTC.Sync)
                newTC.ID = iTC[0]['external_id']
                newTC.FullID = iTC[0]['full_external_id']
                newTC.Address = self._getFullSuitePath(iTC_Details[0]['testsuite_id'])
                newTC.Name = iTC_Details[0]['name']
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
                self.append_Test(newTC)
                tmpResults.append(newTC.FullID)
            return tmpResults
        return []

    def connectTL(self):
        self._testplan()
        self._testbuild()
