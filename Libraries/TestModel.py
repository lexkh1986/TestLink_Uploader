class Test(object):
    def __init__(self, project, testplan, testbuild):
        self.PROJECT_NAME = project
        self.TESTPLAN_NAME = testplan
        self.TESTBUILD_NAME = testbuild
        self.TESTS = []

    def append_Test(self, TestCase):
        self.TESTS.append(TestCase)

    def toDict(self):
        return [iTC.toDict() for iTC in self.TESTS]

class TestCase(object):
    def __init__(self):
        self.ID = None
        self.FullID = None
        self.Name = None
        self.Author = None
        self.Address = None
        self.Summary = None
        self.Steps = None
        self.Priority = None
        self.Exectype = None
        self.Sync = False
        self.WbIndex = None

        self.Result = None
        self.TestedBy = None
        self.Duration = None
        self.Note = None

    def toDict(self):
        return self.__dict__

##test = Test()
##testcase1 = TestCase('Test1', 'Name1', 'Summary1', 'Medium', 'Manual')
##testcase2 = TestCase('Test2', 'Name2', 'Summary2', 'High', 'Automated')
##test.append_Test(testcase1)
##test.append_Test(testcase2)
