import re
from HTMLParser import HTMLParser

TAG_RE = re.compile(r'<[^>]+>')

def dict_getkey(dict_, value):
    return next((key for key, val in dict_.items() if val == value), None)

def getVarFromFile(filename):
    import imp
    with open(filename) as f:
        global data
        data = imp.new_module('data')
        exec(f.read(), data.__dict__)
    return data

def parse_summary(string, isReversed=False):
    if string is None: string = ''
    if type(string) == list: string = str(string)
    if not isReversed:
        val = string.split('''\n''')
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
