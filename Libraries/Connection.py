from TestModel import *
from testlink import *

class Connection(Test):
    STATUS = {'p':'PASS', 'f':'FAIL', 'n':'NOT RUN', 'b':'BLOCK'}
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

    def pullTestCases(self):
        iTC_List = self.CONN.getTestCasesForTestPlan(testplanid = self.TESTPLAN_ID,
                                                     buildid = self.TESTBUILD_ID,
                                                     details = 'simple').values()
        for iTC in iTC_List:
            iTC_Details = self._getTestCase_byID(iTC[0]['full_external_id'])
            if iTC[0]['external_id'] not in self.TESTS.keys():
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
                self.append_Test(newTC)
