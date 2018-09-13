from TestModel import *
from Template import *
from testlink import *
from Misc import *

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
        self._testplan()
        self._testbuild()

    def _project(self):
        for elem in self.CONN.getProjects():
            if elem['name'] == self.PROJECT_NAME:
                self.PROJECT_ID = elem['id']
                self.PROJECT_PREFIX = elem['prefix']
                return
        raise Exception('Project name not found: %s' % self.PROJECT_NAME)

    def _testplan(self):
        self.TESTPLAN_ID = self.CONN.getTestPlanByName(testprojectname = self.PROJECT_NAME,
                                                       testplanname = self.TESTPLAN_NAME)[0]['id']

    def _testbuild(self):
        iBuilds = self.CONN.getBuildsForTestPlan(self.TESTPLAN_ID)
        for i in iBuilds:
            if self.TESTBUILD_NAME == i['name']:
                self.TESTBUILD_ID = i['id']
                return
        raise Exception('Test build not found: %s' % self.TESTBUILD_NAME)

    def _getTestCase_byID(self, full_external_id):
        return self.CONN.getTestCase(testcaseexternalid = full_external_id)

    def _getPath(self, full_path, delimeter='/'):
        tmpPath = full_path.split(delimeter)
        tmpRefID, tmpRefName = None, None
        tmpRoot = self.CONN.getFirstLevelTestSuitesForTestProject(self.PROJECT_ID)

        # Get first child
        for fc in tmpRoot:
            if fc['name'] == tmpPath[0]:
                tmpRefID, tmpRefName = fc['id'], fc['name']
        if tmpRefID is None: raise Exception('Could not found "%s" in full path: %s' % (tmpPath[0], full_path))

        # Get child loop
        for idx, node in enumerate(tmpPath):
            if idx == 0: continue
            tmpChilds = self.CONN.getTestSuitesForTestSuite(tmpRefID)
            if not tmpChilds: raise Exception('Could not found "%s" in full path: %s' % (node, full_path))
            tmpChilds = tmpChilds.values() if tmpChilds.get('name', None) is None else [tmpChilds]
            for n in tmpChilds:
                if n['name'] == node:
                    tmpRefID, tmpRefName = n['id'], n['name']
                    isFound = True
        return tmpRefID

    def pushTestCase(self, iTC_):
        if iTC_.FullID in ('', None):
            try:
                rs = self.CONN.createTestCase(testcasename = iTC_.Name,
                                              testprojectid = self.PROJECT_ID,
                                              testsuiteid = self._getPath(iTC_.Address, self.DELIMETER),
                                              authorlogin = iTC_.Author,
                                              summary = parse_summary(iTC_.Summary),
                                              steps = parse_summary(iTC_.Steps),
                                              importance = dict_getkey(self.IMPORTANCE, iTC_.Priority),
                                              executiontype = dict_getkey(self.EXECUTIONTYPE, iTC_.Exectype))
                iTC_.ID = rs[0]['additionalInfo']['external_id']
                if iTC_.ID is not None:
                    iTC_.FullID = '%s-%s' % (self.PROJECT_PREFIX, iTC_.ID)
                    iTC_.SyncStatus = 3
            except Exception, err: print 'Failed to create TestCase: %s\n%s' % (iTC_.Name, err)
        else:
            try:
                self.CONN.updateTestCase(testcaseexternalid = iTC_.FullID,
                                         testcasename = iTC_.Name,
                                         user = iTC_.Author,
                                         summary = parse_summary(iTC_.Summary),
                                         steps = parse_summary(iTC_.Steps),
                                         importance = dict_getkey(self.IMPORTANCE, iTC_.Priority),
                                         executiontype = dict_getkey(self.EXECUTIONTYPE, iTC_.Exectype))
                iTC_.SyncStatus = 4
            except Exception, err: print 'Failed to modifiy TestCase: %s\n%s' % (iTC_.Name, err)

    def pullTestCases(self):
        iTemplate = Template()
        iStart = iTemplate.LOC_DETAILS.get('r') + 1
        iConsumedLoc = [elem.toDict()['WbIndex'] for elem in self.TESTS]
        iOpenLoc = [i for i in range(iStart, iTemplate.MAX_ROW_SUPPORT) if i not in iConsumedLoc]
        iFullID = [iTC.FullID for iTC in self.TESTS]
        iTC_List = self.CONN.getTestCasesForTestPlan(testplanid = self.TESTPLAN_ID,
                                                     buildid = self.TESTBUILD_ID,
                                                     details = 'simple')
        if not iTC_List:
            return False

        iTC_List = iTC_List.values()
        for iTC in iTC_List:
            if iTC[0]['full_external_id'] in iFullID: continue
            iTC_Details = self._getTestCase_byID(iTC[0]['full_external_id'])
            newTC = TestCase()
            newTC.Sync = dict_getkey(self.SYNC, newTC.Sync)
            newTC.SyncStatus = 1
            newTC.ID = iTC[0]['external_id']
            newTC.FullID = iTC[0]['full_external_id']
            newTC.Name = iTC_Details[0]['name']
            newTC.Summary = iTC_Details[0]['summary']
            newTC.Steps = iTC_Details[0]['steps']
            newTC.Author = iTC_Details[0]['author_login'].lower()
            newTC.Version = int(iTC_Details[0]['version'])
            newTC.Priority = self.IMPORTANCE.get(int(iTC_Details[0]['importance']))
            newTC.Exectype = self.EXECUTIONTYPE.get(int(iTC_Details[0]['execution_type']))
            if iTC[0]['exec_status'] != 'n':
                newTC.Result = self.STATUS.get(iTC[0]['exec_status'])
                newTC.Duration = iTC[0]['execution_duration']
            iTakenLoc = iOpenLoc.pop(0)
            newTC.WbIndex = iTakenLoc
            self.append_Test(newTC)
        return True
