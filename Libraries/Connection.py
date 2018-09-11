from testlink import *

class Connection(object):
    def __init__(self, devkey, project):
        self.SERVER_URL = 'http://testlink.nexcel.vn/lib/api/xmlrpc/v1/xmlrpc.php'
        self.DEVKEY = devkey
        self.CONN = TestLinkHelper(self.SERVER_URL, self.DEVKEY).connect(TestlinkAPIGeneric)

        for elem in self.CONN.getProjects():
            if elem['name'] == project:
                self.PROJECT_ID = elem['id']
                self.PROJECT_PREFIX = elem['prefix']
                self.PROJECT_NAME = project
                print 'Connected to TestLink server with project: %s(ID: %s)' % (project, self.PROJECT_ID)
                return
        raise Exception('Could not found project name %s' % project)

    def conn(self):
        return self.CONN
