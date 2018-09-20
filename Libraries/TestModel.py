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

    def get_byFullID(self, fullID):
        for iTC in self.TESTS:
            if iTC.FullID == fullID:
                return iTC
        return None

    def pop_byFullID(self, fullID):
        for idx, iTC in enumerate(self.TESTS):
            if iTC.FullID == fullID:
                del self.TESTS[idx]

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
        self.Owner = None
        self.Version = 1
        self.Sync = 0
        self.WbIndex = None
        self.fmtCode = None

        self.Result = None
        self.Duration = None
        self.Note = None

    def toDict(self):
        return self.__dict__
