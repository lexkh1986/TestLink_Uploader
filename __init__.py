import os, sys
sys.path.append('Libraries')
from Workbook import Workbook

if __name__ == "__main__":
    if set(sys.argv) & set(['-p','--pull']):
        wb = Workbook()
        wb.loadWorkbook()
        wb.loadTestCases()
        print 'Test Project: %s' % wb.INFO.PROJECT_NAME
        print 'Test Plan: %s' % wb.INFO.TESTPLAN_NAME
        print 'Test Build: %s' % wb.INFO.TESTBUILD_NAME
        print '\nDo commands:'
        wb.pullTestCases()

    if set(sys.argv) & set(['-s','--sync']):
        wb = Workbook()
        wb.loadWorkbook()
        wb.loadTestCases()
        print 'Test Project: %s' % wb.INFO.PROJECT_NAME
        print 'Test Plan: %s' % wb.INFO.TESTPLAN_NAME
        print 'Test Build: %s' % wb.INFO.TESTBUILD_NAME
        print '\nDo commands:'
        wb.pushTestCases()

    if set(sys.argv) & set(['-sr','--syncResult']):
        wb = Workbook()
        wb.loadWorkbook()
        wb.loadTestCases(False)
        print 'Test Project: %s' % wb.INFO.PROJECT_NAME
        print 'Test Plan: %s' % wb.INFO.TESTPLAN_NAME
        print 'Test Build: %s' % wb.INFO.TESTBUILD_NAME
        print '\nDo commands:'
        wb.pushResults()
