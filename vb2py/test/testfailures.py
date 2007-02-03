"""Tests that we know fail but are not within the scope of v0.2"""


from unittest import *
from testframework import *

# << Failing tests >> (1 of 2)
# Simple function with ByRef argument which is changed
tests.append(("""
Function _square(x, y)
    _square = x*x
    y = y + 1
End Function
b = 0
a = _square(10, b)
""", {"a" : 100, "b" : 1}))
# << Failing tests >> (2 of 2)
# Optional arguments
tests.append(("""
Sub _Change(ByVal x, ByRef y)
    x = x + 1
    y = y + 1
End Sub
a = 0
b = 0
_Change a, b
""", {"a" : 0, "b" : 1}))
# -- end -- << Failing tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
