from unittest import *
from testframework import *

# << Const tests >>
# Simple test
tests.append(("""
Const a = 10
Const b = 20
Const c = "hello"
""", {"a":10, "b":20, "c":"hello"}))

# Simple test woth scope
tests.append(("""
Private Const a = 10
Private Const b = 20
Private Const c = "hello"
""", {"a":10, "b":20, "c":"hello"}))

tests.append(("""
Public Const a = 10
Public Const b = 20
Public Const c = "hello"
""", {"a":10, "b":20, "c":"hello"}))

# Simple test with hex numbers
tests.append(("""
Const a = &HFF
Const b = &HA
""", {"a":255, "b":10, }))

# Three on a line
tests.append(("""
Const _a = "one", _b = "two", _c = "three"
d = _a & _b & _c
""", {"d" : "onetwothree"}))
# -- end -- << Const tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
