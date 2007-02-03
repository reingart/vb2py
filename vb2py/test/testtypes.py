from unittest import *
from testframework import *

# << Type tests >> (1 of 2)
# Simple test
tests.append(("""
Type _Point
    X As Single
    Y As Single
End Type
Dim _a As _Point
_a.X = 10
_a.Y = 20
b = _a.X + _a.Y
""", {"b" : 30}))

# Nested Types
tests.append(("""
Type _Point
    X As Single
    Y As Single
End Type
Type _Line
    P1 As _Point
    P2 As _Point
End Type
Dim _a As _Line
_a.P1.X = 10
_a.P2.X = 20
_a.P1.Y = 1
_a.P2.Y = 2
b = _a.P1.X + _a.P1.Y
c = _a.P2.X + _a.P2.Y
""", {"b" : 11, "c" : 22}))

# Empty type
tests.append(("""
Type _Point
End Type
Dim _a As _Point
""", {}))
# << Type tests >> (2 of 2)
# Arrays of a type
tests.append(("""
Type _Point
    X As Single
    Y As Single
End Type

Dim _p(5) As _Point

tx = 0
ty = 0
For _i = 1 To 5
    _p(_i).X = _i
    _p(_i).Y = 2*_i
    tx = tx + _p(_i).X
    ty = ty + _p(_i).Y
Next _i
""", {"tx" : 15, "ty" : 30}))
# -- end -- << Type tests >>

import vb2py.vbparser
vb2py.vbparser.log.setLevel(0) # Don't print all logging stuff
TestClass = addTestsTo(BasicTest, tests)

if __name__ == "__main__":
    main()
