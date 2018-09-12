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
    if not isReversed:
        val = string.split('''\n''')
        for i, v in enumerate(val):
            val[i] = '%s<br/>' % v
            val[i] = val[i].replace('Step:','<strong>&emsp;Step:</strong>')
            val[i] = val[i].replace('Checkpoint:','<strong>&emsp;Checkpoint:</strong>')
            val[i] = val[i].replace('Verify point:','<strong>&emsp;Verify point:</strong>')
        return ''.join(val)
    else:
        val = string.split('<br/>')
        if '\n' in val: val.remove('\n')
        for i, v in enumerate(val):
            val[i] = val[i].strip(' \t\r')
            val[i] = val[i].replace('<strong>&emsp;Step:</strong>', 'Step:')
            val[i] = val[i].replace('<strong>&emsp;Checkpoint:</strong>', 'Checkpoint:')
            val[i] = val[i].replace('<strong>&emsp;Verify point:</strong>', 'Verify point:')
        return '\n'.join(val)
