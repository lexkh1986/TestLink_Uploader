from Connection import *
from TestModel import *
from Misc import *
import os

class Workbook(Test):
    CONFIG_PATH = '%s/../setting.txt' % os.path.dirname(__file__)
    
    def __init__(self):
        super(Workbook, self).__init__()

    def _setConfig(self):
        if os.path.isfile(self.CONFIG_PATH):
            iConfig = getVarFromFile(self.CONFIG_PATH)
            self.FILE = iConfig.FILE_PATH
            self.USER = iConfig.USERNAME
            self.DEVKEY = iConfig.DEVKEY

    def build(self):
        self._setConfig()

a = Workbook()
a.build()
