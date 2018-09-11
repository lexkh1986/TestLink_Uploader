class Test(object):
    def __init__(self):
        self.PROJECT_NAME = None
        self.TESTPLAN_NAME = None
        self.TESTBUILD_NAME = None
        self.TESTS = {}

    def append_Test(self, TestCase):
        self.TESTS.update({len(self.TESTS.keys())+1:TestCase})

    def get_Test(self, index):
        return self.TESTS.get(index).toDict()

    def set_Project(self, name):
        self.PROJECT_NAME = name

    def set_TestPlan(self, name):
        self.TESTPLAN_NAME = name

    def set_TestBuild(self, name):
        self.TESTBUILD_NAME = name

class TestCase(object):
    def __init__(self, id, name, address, priority, exectype):
        self.ID = id
        self.Name = name
        self.Address = address
        self.Steps = None
        self.Priority = priority
        self.Exectype = exectype
        self.Sync = False

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
##print test.get_Test(1)
##print test.get_Test(2)
