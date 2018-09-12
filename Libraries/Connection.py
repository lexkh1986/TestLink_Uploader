from TestModel import *
from Template import *
from testlink import *
from Misc import *

class Connection(Test):
    STATUS = {'p':'PASS', 'f':'FAIL', 'n':'NOT RUN', 'b':'BLOCK'}
    STATE = {2:'READFORREVIEW', 4:'REWORK', 3:'FINAL', 1:'DRAFT'}
    IMPORTANCE = {3:'HIGH', 2:'MEDIUM', 1:'LOW'}
    EXECUTIONTYPE = {1:'MANUAL', 2:'AUTOMATED'}

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

    def pushTestCases(self):
        for iTC in self.TESTS:
            if iTC.Sync:
                if iTC.FullID in ('', None):
                    rs = self.CONN.createTestCase(testcasename = iTC.Name,
                                                  testprojectid = self.PROJECT_ID,
                                                  testsuiteid = self._getPath(iTC.Address),
                                                  authorlogin = iTC.Author,
                                                  summary = parse_summary(iTC.Summary),
                                                  steps = parse_summary(iTC.Steps),
                                                  importance = dict_getkey(self.IMPORTANCE, iTC.Priority),
                                                  executiontype = dict_getkey(self.EXECUTIONTYPE, iTC.Exectype))
                    iTC.FullID = '%s-%s' % (self.PROJECT_PREFIX, rs[0]['additionalInfo']['external_id'])
                else:
                    try:
                        self.CONN.updateTestCase(testcaseexternalid = iTC.FullID,
                                                 testcasename = iTC.Name,
                                                 user = iTC.Author,
                                                 summary = parse_summary(iTC.Summary),
                                                 steps = parse_summary(iTC.Steps),
                                                 importance = dict_getkey(self.IMPORTANCE, iTC.Priority),
                                                 executiontype = dict_getkey(self.EXECUTIONTYPE, iTC.Exectype))
                    except: pass

    def pullTestCases(self):
        iTemplate = Template()
        iStart = iTemplate.LOC_DETAILS.get('r') + 1
        iConsumedLoc = [elem.toDict()['WbIndex'] for elem in self.TESTS]
        iOpenLoc = [i for i in range(iStart, iTemplate.MAX_ROW_SUPPORT) if i not in iConsumedLoc]
        iFullID = [iTC.FullID for iTC in self.TESTS]
        iTC_List = self.CONN.getTestCasesForTestPlan(testplanid = self.TESTPLAN_ID,
                                                     buildid = self.TESTBUILD_ID,
                                                     details = 'simple').values()

        for iTC in iTC_List:
            if iTC[0]['full_external_id'] in iFullID: continue
            iTC_Details = self._getTestCase_byID(iTC[0]['full_external_id'])
            newTC = TestCase()
            newTC.ID = iTC[0]['external_id']
            newTC.FullID = iTC[0]['full_external_id']
            newTC.Name = iTC_Details[0]['name']
            newTC.Summary = iTC_Details[0]['summary']
            newTC.Steps = iTC_Details[0]['steps']
            newTC.Author = iTC_Details[0]['author_login']
            newTC.Priority = self.IMPORTANCE.get(int(iTC_Details[0]['importance']))
            newTC.Exectype = self.EXECUTIONTYPE.get(int(iTC_Details[0]['execution_type']))
            if iTC[0]['exec_status'] != 'n':
                newTC.Result = self.STATUS.get(iTC[0]['exec_status'])
                newTC.Duration = iTC[0]['execution_duration']
            iTakenLoc = iOpenLoc.pop(0)
            newTC.WbIndex = iTakenLoc
            self.append_Test(newTC)
