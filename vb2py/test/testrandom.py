# Created by Leo from: C:\Development\Python22\Lib\site-packages\vb2py\vb2py.leo

from unittest import *
from testframework import *

# << Random tests >>
# Rnd produces different numbers
tests.append(("""
_a = Rnd
_b = Rnd
_c = Rnd
a = _a = _b
b = _b = _c
""", {"a":0, "b":0}))


# Rnd with 0 produces the last value
tests.append(("""
_a = Rnd
_b = Rnd(0)
_c = Rnd
a = _a = _b
b = _b = _c
""", {"a":1, "b":0}))

# Rnd with -ve produces consistent sequence
tests.append(("""
_a1 = Rnd(-100)
_b1 = Rnd
_c1 = Rnd
_a2 = Rnd(-100)
_b2 = Rnd
_c2 = Rnd
a = _a1 = _a1
b = _b1 = _b1
c = _c1 = _c1
""", {"a":1, "b":1, "c":1}))


# Randomize breaks a sequence
tests.append(("""
_a1 = Rnd(-100)
_b1 = Rnd
_c1 = Rnd
_a2 = Rnd(-100)
Randomize
_b2 = Rnd
_c2 = Rnd
a = _a1 = _a2
b = _b1 = _b2
c = _c1 = _c2
""", {"a":1, "b":0, "c":0}))

# Randomize breaks a sequence
tests.append(("""
_a1 = Rnd(-100)
_b1 = Rnd
_c1 = Rnd
_a2 = Rnd(-100)
Randomize 25
_b2 = Rnd
_c2 = Rnd
a = _a1 = _a2
b = _b1 = _b2
c = _c1 = _c2
""", {"a":1, "b":0, "c":0}))
# -- end -- << Random tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
	main()
