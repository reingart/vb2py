from testframework import *
import sys

# << External tests >>
# Simple test
if sys.platform in ('win32', 'win64'):
    tests.append(("""
    Dim _a As Object
    Set _a = CreateObject("Excel.Application")
    b = _a.Name
    """, {"b":"Microsoft Excel"}))
# -- end -- << External tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
