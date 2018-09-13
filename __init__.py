import os, sys
sys.path.append('Libraries')
from Workbook import Workbook

if __name__ == "__main__":
    if set(sys.argv) & set(['-p','--pull']):
        wb = Workbook()
        wb._loadWorkbook()
        print 'Test Project: %s' % wb.INFO.PROJECT_NAME
        print 'Test Plan: %s' % wb.INFO.TESTPLAN_NAME
        print 'Test Build: %s' % wb.INFO.TESTBUILD_NAME
        wb._loadTestCases()
        wb._pullTestCases()

    if set(sys.argv) & set(['-s','--sync']):
        wb = Workbook()
        wb._loadWorkbook()
        print 'Test Project: %s' % wb.INFO.PROJECT_NAME
        print 'Test Plan: %s' % wb.INFO.TESTPLAN_NAME
        print 'Test Build: %s' % wb.INFO.TESTBUILD_NAME
        wb._loadTestCases()
        wb._pushTestCases()
