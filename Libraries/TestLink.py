from HTMLParser import HTMLParser
from testlink import *
from numpy import *
import csv
import re

class TestLink():
    IMPORTANCE = {'HIGH':3, 'MEDIUM':2, 'LOW':1}
    EXECUTIONTYPE = {'MANUAL':1, 'AUTOMATED':2}
    STATE = {'READFORREVIEW':2, 'REWORK':4, 'FINAL':3, 'DRAFT':1}
    BOOLEAN = {'TRUE':1, 'FALSE':0}
    RESULT = {'PASS':'p', 'FAIL':'f'}
 
    def __init__(self):
        self.SERVER_URL = 'http://testlink.nexcel.vn/lib/api/xmlrpc/v1/xmlrpc.php'

    def _project(self, project_name):
        for elem in self.CONN.getProjects():
            if elem['name'] == project_name:
                self.PROJECT_ID = elem['id']
                self.PROJECT_PREFIX = elem['prefix']
                self.PROJECT_NAME = elem['name']
                return
        raise Exception('Could not found project name %s' % project_name)

    def _root(self):
        return array(self.CONN.getFirstLevelTestSuitesForTestProject(self.PROJECT_ID))

    def _getPath(self, full_path, delimeter='/'):
        tmpPath = full_path.split(delimeter)
        tmpRefID, tmpRefName = None, None

        # Get first child
        for fc in self._root():
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

    def _createTestCase(self, tc_name, full_path, tc_author, tc_summary='', tc_step='', tc_status='Draft', tc_importance='Medium', tc_executiontype='Manual'):
        rs = self.CONN.createTestCase(testcasename = tc_name,
                                        testprojectid = self.PROJECT_ID,
                                        testsuiteid = self._getPath(full_path),
                                        authorlogin = tc_author,
                                        summary = tc_summary,
                                        steps = tc_step,
                                        status = self.STATE.get(tc_status.upper()),
                                        importance = self.IMPORTANCE.get(tc_importance.upper()),
                                        executiontype = self.EXECUTIONTYPE.get(tc_executiontype.upper()))[0]
        return {'external_id':rs['additionalInfo']['external_id'], 'status':rs['message']}

    def _updateTestCase(self, external_id, tc_name, tc_author, tc_summary='', tc_step='', tc_status='Draft', tc_importance='Medium', tc_executiontype='Manual', tc_version=1):
        try:
            self.CONN.updateTestCase(testcaseexternalid = self.PROJECT_PREFIX + '-' + str(external_id),
                                   testcasename = tc_name,
                                   user = tc_author,
                                   summary = tc_summary,
                                   steps = tc_step,
                                   status = self.STATE.get(tc_status.upper()),
                                   importance = self.IMPORTANCE.get(tc_importance.upper()),
                                   executiontype = self.EXECUTIONTYPE.get(tc_executiontype.upper()))
        except: pass
        return None

    def _reportTestResult(self, testplan, tc_external_id, tc_status, tc_build=None, tc_note=None, exec_user=None, isOverwrite=False):
        return self.CONN.reportTCResult(testplanid = testplan,
                                        testcaseexternalid = self.PROJECT_PREFIX + '-' + str(tc_external_id),
                                        status = tc_status,
                                        buildname = tc_build,
                                        notes = tc_note,
                                        user = exec_user,
                                        overwrite = isOverwrite)

    def _getTestCase_byID(self, external_id):
        return self.CONN.getTestCase(testcaseexternalid = self.PROJECT_PREFIX + '-' + str(external_id))

    def _getTestCase_byName(self, tc_name):
        try:
            return self.CONN.getTestCaseIDByName(testcasename = tc_name, testprojectname = self.PROJECT_NAME)
        except:
            return []

    def _getPlan(self, testplanname):
        return self.CONN.getTestPlanByName(testprojectname = self.PROJECT_NAME,
                                           testplanname = testplanname)

    def _getTCList_byPlan(self, testplanname):
        return self.CONN.getTestCasesForTestPlan(testplanid = self._getPlan(testplanname)[0]['id'])

    def _getTCList_bySuite(self, fullpath, delimeter='/', isDeep=False, isDetails='simple'):
        suiteID = self._getPath(fullpath, delimeter)
        return self.CONN.getTestCasesForTestSuite(testsuiteid = suiteID, deep = isDeep, details = isDetails)

    @staticmethod
    def _getkey(dictt, keyword):
        for key, val in dictt.items():
            if val == keyword:
                return key.capitalize()

    @staticmethod
    def _html_decode(string):
        return string
        #return HTMLParser().unescape(string)

    @staticmethod
    def connectTestLink(userkey, project_name):
        """
        Connect to TestLink server project and return the connection
        """
        iConn = TestLink()
        iConn.DEVKEY = userkey
        iConn.CONN = TestLinkHelper(iConn.SERVER_URL, iConn.DEVKEY).connect(TestlinkAPIGeneric)
        iConn._project(project_name)
        print('Connected to TestLink project: %s (server version %s)' % (project_name, iConn.CONN.testLinkVersion()))

        return iConn

    @staticmethod
    def markResult_all(tlConn_, testplan, status, isOverwrite=False):
        """
        Mark all test results for a testplan
        """
        iPlanID = tlConn_._getPlan(testplan)[0]['id']
        iTCLists = [(elem[0]['external_id'], elem[0]['tcase_name']) for elem in tlConn_._getTCList_byPlan(testplan).values()]
        iBuildLists = [(elem['id'], elem['name']) for elem in tlConn_.CONN.getBuildsForTestPlan(iPlanID)]
        for tb in iBuildLists:
            for i, tc in enumerate(iTCLists):
                tlConn_.CONN.reportTCResult(testplanid = iPlanID,
                                            testcaseexternalid = tlConn_.PROJECT_PREFIX + '-' + str(tc[0]),
                                            buildid = tb[0],
                                            status = tlConn_.RESULT.get(status.upper()))
                print 'Updated test result %s for TestID: %s - %s' % (status, tc[0], tc[1])

    @staticmethod
    def exportTestCase_csv(tlConn_, suitepath, suitepath_delimeter, filepath):
        """
        Export all testcases in a testsuite (i.e. suite name "child" in sample) to csv file\n
        Path format sample: grandparent/parent/child
        """
        iData = tlConn_._getTCList_bySuite(suitepath, suitepath_delimeter, True, 'only_id')
        iDataDetails = [tlConn_.CONN.getTestCase(testcaseid = x) for x in iData]
        iResult = []
        for tc in iDataDetails:
            row = {'id':tc[0]['tc_external_id'],
                   'testcase_name':tc[0]['name'],
                   'fullpath':'',
                   'author':tc[0]['author_login'],
                   'summary':tlConn_._html_decode(tc[0]['summary']),
                   'status':tlConn_._getkey(tlConn_.STATE, int(tc[0]['status'])),
                   'importance':tlConn_._getkey(tlConn_.IMPORTANCE, int(tc[0]['importance'])),
                   'exectype':tlConn_._getkey(tlConn_.EXECUTIONTYPE, int(tc[0]['execution_type'])),
                   'execduration':tc[0]['estimated_exec_duration'],
                   'version':tc[0]['version']}
            iResult.append(row)

        keys = ['id', 'testcase_name', 'fullpath', 'author', 'summary', 'status', 'importance', 'exectype', 'execduration', 'version']
        with open(filepath, 'wb') as of:
            dict_writer = csv.DictWriter(of, keys)
            dict_writer.writeheader()
            dict_writer.writerows(iResult)

    @staticmethod
    def _parse_summary(string):
        val = string.split('''\n''')
        for i, v in enumerate(val):
            val[i] = '%s<br/>' % v
            val[i] = val[i].replace('Step:','<strong>&emsp;Step:</strong>')
            val[i] = val[i].replace('Checkpoint:','<strong>&emsp;Checkpoint:</strong>')
            val[i] = val[i].replace('Verify point:','<strong>&emsp;Verify point:</strong>')
        completedStr = ''.join(val)
        return completedStr

    @staticmethod
    def uploadTestCase_csv(tlConn_, filepath):
        """
        Upload test cases to TestLink from a csv file\n
        Format columns (name, fullpath, author are must have):\n
        Mark import_flag as FALSE for each line if you want to skip importing them without deleting row from the csv
            import_flag | id    | testcase_name | fullpath | author | summary | status | importance | exectype  | execduration | version\n
            ----------- | ----- | ------------- | -------- | ------ | ------- | ------ | ---------- | --------- | ------------ | -------\n
            TRUE        | 123   | Test1         | S1/S2/S3 | lex.kh | Test1s  | Final  | High       | Manual    | 1m           | 1      \n
            FALSE       |       | Test2         | S1/S2/S4 | lex.kh | Test2s  | Draft  | Low        | Automated | 2m           | 2      \n
        """
        iData = []
        with open(filepath) as f:
            reader = csv.reader(f, skipinitialspace=True)
            header = next(reader)
            iData = [dict(zip(header, row)) for row in reader]

        # Validate name
        iInvalidName = []
        for tc in iData:
            tmpFoundList = tlConn_._getTestCase_byName(tc['testcase_name'])
            if not tmpFoundList:
                continue
            else:
                tmpExterIDList = [i.get('tc_external_id') for i in tmpFoundList]
                if tc['id'] not in ('', None) and tc['id'] in tmpExterIDList:
                    continue
            iInvalidName.append(tc['testcase_name'])

        if iInvalidName:
            raise Exception('Some names already exist in project. Please recheck to make sure not duplicate:\n' + '\n'.join(iInvalidName))
        
        # Upload row by row
        for tc in iData:
            tc['summary'] = tlConn_._parse_summary(tc['summary'])
            if not bool(tlConn_.BOOLEAN.get(tc['import_flag'].upper())):
                continue
            if tc['id'] in ('', None):
                result = tlConn_._createTestCase(tc_name = tc['testcase_name'],
                                                 full_path = tc['fullpath'],
                                                 tc_author = tc['author'],
                                                 tc_summary = tc['summary'],
                                                 tc_status = tc['status'],
                                                 tc_importance = tc['importance'],
                                                 tc_executiontype = tc['exectype'])
                print 'Imported new: %s - %s' % (result['external_id'], tlConn_._getTestCase_byID(result['external_id'])[0]['name'])
            else:
                result = tlConn_._updateTestCase(external_id = tc['id'],
                                                 tc_name = tc['testcase_name'],
                                                 tc_author = tc['author'],
                                                 tc_summary = tc['summary'],
                                                 tc_status = tc['status'],
                                                 tc_importance = tc['importance'],
                                                 tc_executiontype = tc['exectype'],
                                                 tc_version = tc['version'])
                print 'Updated: %s - %s' % (tc['id'], tc['testcase_name'])
