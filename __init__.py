import os, sys
sys.path.append('Libraries')
from Process import Workbook

if __name__ == "__main__":
    if set(sys.argv) & set(['-p','--pull']):
        wb = Workbook()
        wb.loadWorkbook()
        wb.connect()
        wb.loadTestCases()
        print('\nImportant Note: Each cell in excel file can only support not more than 32767 chars. Any chars more than 32767 in step/summary will be cut off.')
        print('\nDo commands:')
        wb.pullTestCases()

    if set(sys.argv) & set(['-s','--sync']):
        wb = Workbook()
        wb.loadWorkbook()
        wb.connect()
        wb.loadTestCases()
        print('\nDo commands:')
        wb.pushTestCases()

    if set(sys.argv) & set(['-sr','--syncResult']):
        wb = Workbook()
        wb.loadWorkbook()
        wb.connect()
        wb.loadTestCases()
        print('\nDo commands:')
        wb.pushResults()

    if set(sys.argv) & set(['-cr','--createTestPlan']):
        wb = Workbook()
        wb.loadWorkbook()
        print('\nDo commands:')
        wb.createTestplan()

    if set(sys.argv) & set(['-at','--addToTestPlan']):
        wb = Workbook()
        wb.loadWorkbook()
        wb.connect()
        wb.loadTestCases()
        print('\nDo commands:')
        wb.addToTestPlan()
