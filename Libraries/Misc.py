def dict_getkey(dict_, value):
    return next((key for key, val in dict_.items() if val == value), None)

def getVarFromFile(filename):
    import imp
    with open(filename) as f:
        global data
        data = imp.new_module('data')
        exec(f.read(), data.__dict__)
    return data
