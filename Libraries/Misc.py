import re
import os, sys
from HTMLParser import HTMLParser

TAG_RE = re.compile(r'<[^>]+>')

def dict_getkey(dict_, value, default=None):
    return next((key for key, val in dict_.items() if val == value), default)

def getVarFromFile(filename):
    import imp
    with open(filename) as f:
        global data
        data = imp.new_module('data')
        exec(f.read(), data.__dict__)
    return data

def parse_steps(steps, isReversed=False):
    if isReversed:
        if steps not in (None, ''):
            tmpStr = ['Step%s: %s\nExpectation: %s\n' % (s.get('step_number'),
                                                     s.get('actions'),
                                                     s.get('expected_results'))\
                      for s in steps]
            return '\n'.join(tmpStr)
        return steps
    else:
        if steps not in (None, ''):
            tmpStr = steps.encode('ascii',errors='ignore').split('\n\n')
            tmpSteps = []
            for s in tmpStr:
                tmpSt = s.split('\n')
                tmpNo = tmpSt[0].split(': ')
                tmpValidated = {}
                tmpValidated.update({'step_number':int(tmpNo[0].strip('Step')),
                                     'actions':tmpNo[1],
                                     'expected_results':tmpSt[1].strip('. Expectation: ')})
                tmpSteps.append(tmpValidated)
            return tmpSteps
        return steps

def parse_summary(string, isReversed=False):
    if string is None: string = ''
    if type(string) == list: string = str(string)
    if not isReversed:
        val = string.split('\n')
        for i, v in enumerate(val):
            val[i] = '%s<br/>' % v
            val[i] = val[i].replace('Step:','<strong>&emsp;Step:</strong>')
            val[i] = val[i].replace('Checkpoint:','<strong>&emsp;Checkpoint:</strong>')
            val[i] = val[i].replace('Verify point:','<strong>&emsp;Verify point:</strong>')
        return ''.join(val)
    else:
        ps = HTMLParser()
        val = string.encode('ascii',errors='ignore').split('<br/>')
        for i, v in enumerate(val):
            val[i] = re.sub(r'\n\t', '', val[i])
            val[i] = remove_tags(val[i])
            val[i] = ps.unescape(val[i])
        val = filter(None, val)
        return '\n'.join(val).strip()

def remove_tags(text):
    return TAG_RE.sub('', text)

def isReadonly(filepath):
    if os.path.exists(filepath):
        try:
            os.rename(filepath, filepath)
        except OSError as e:
            print 'Permission denied: %s\nPlease close your workbook and re-run task again.' % filepath
            sys.exit(1)
            
def rem_empty(string):
    return string.replace('\n','').replace('\t','').replace(' ','').encode('ascii','ignore')
