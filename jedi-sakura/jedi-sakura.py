# coding: utf-8

import sys
import jedi

line = int(sys.argv[1])
column = int(sys.argv[2])
ipath = sys.argv[3]
epath = sys.argv[4]
#print line
#print column
#print ipath
#print epath

script = jedi.Script(source = open(epath).read(), line = line, column = column, path = ipath)

# drop _prop, __prop
cmpls = [n for n in script.completions() if not n.name.startswith('_')]
# sort by type, name
#   ... makes no sense
#cmpls = sorted(cmpls, key=lambda c: (c.type, c.name))


for cmpl in cmpls:
    if cmpl.type in [ 'function', 'class' ]:
        # drop *args, **kwargs, ...
        params = ', '.join(p.description for p in cmpl.params if not p.description.startswith('*') and not p.description.startswith('...'))
        print "%s(%s)" % (cmpl.name, params)
    else:
        print cmpl.name
