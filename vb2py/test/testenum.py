# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << Enum tests >>
# Simple test
tests.append(("""
Enum thing
	_one
	_two
	_three
	_four
End Enum

a = _one
b = _two
c = _three
d = _four
""", {"a":0, "b":1, "c":2, "d":3}))


# Simple test with values
tests.append(("""
Enum thing
	_one = 1
	_two = 2
	_three = 3
	_four = 4
End Enum

a = _one
b = _two
c = _three
d = _four
""", {"a":1, "b":2, "c":3, "d":4}))
# -- end -- << Enum tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
